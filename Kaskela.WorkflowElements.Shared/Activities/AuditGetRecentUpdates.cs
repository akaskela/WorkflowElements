using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Data;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class AuditGetRecentUpdates : ContributingClasses.AuditBase
    {
        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            this.ValidateColorCode(this.Table_BorderColor.Get(context), "Table - Border Color");
            this.ValidateColorCode(this.Header_BackgroundColor.Get(context), "Header - Background Color");
            this.ValidateColorCode(this.Header_FontColor.Get(context), "Header - Font Color");
            this.ValidateColorCode(this.Row_BackgroundColor.Get(context), "Row - Background Color");
            this.ValidateColorCode(this.Row_FontColor.Get(context), "Row - Font Color");
            this.ValidateColorCode(this.AlternatingRow_BackgroundColor.Get(context), "Alternating Row - Background Color");
            this.ValidateColorCode(this.AlternatingRow_FontColor.Get(context), "Alternating Row - Font Color");

            DataTable table = this.BuildDataTable(context, workflowContext, service);
            this.NumberOfFields.Set(context, table.Rows.Count);
            this.BuildResultsAsHtml(table, context, service);
            this.BuildResultsAsCsv(table, context);
        }

        protected DataTable BuildDataTable(CodeActivityContext context, IWorkflowContext workflowContext, IOrganizationService service)
        {
            TimeZoneSummary timeZone = StaticMethods.CalculateTimeZoneToUse(this.TimeZoneOption.Get(context), workflowContext, service);
            
            DataTable table = new DataTable() { TableName = workflowContext.PrimaryEntityName };
            table.Columns.AddRange(new DataColumn[] { new DataColumn("Date"), new DataColumn("User"), new DataColumn("Attribute"), new DataColumn("Old Value"), new DataColumn("New Value") });
            DateTime oldestUpdate = DateTime.MinValue;

            if (this.Units != null && this.Number != null && this.Number.Get<int>(context) != 0)
            {
                OptionSetValue value = this.Units.Get<OptionSetValue>(context);
                if (value != null)
                {
                    switch (value.Value)
                    {
                        case 222540000:
                            oldestUpdate = DateTime.Now.AddYears(this.Number.Get<int>(context) * -1);
                            break;
                        case 222540001:
                            oldestUpdate = DateTime.Now.AddMonths(this.Number.Get<int>(context) * -1);
                            break;
                        case 222540002:
                            oldestUpdate = DateTime.Now.AddDays(this.Number.Get<int>(context) * -7);
                            break;
                        case 222540003:
                            oldestUpdate = DateTime.Now.AddDays(this.Number.Get<int>(context) * -1);
                            break;
                        case 222540004:
                            oldestUpdate = DateTime.Now.AddHours(this.Number.Get<int>(context) * -1);
                            break;
                        default:
                            oldestUpdate = DateTime.Now.AddMinutes(this.Number.Get<int>(context) * -1);
                            break;
                    }
                }
            }

            int maxUpdates = this.MaxAuditLogs.Get(context) > 100 || this.MaxAuditLogs.Get(context) < 1 ? 100 : this.MaxAuditLogs.Get(context);
            RetrieveRecordChangeHistoryRequest request = new RetrieveRecordChangeHistoryRequest()
            {
                Target = new EntityReference(workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId),
                PagingInfo = new PagingInfo() { Count = maxUpdates, PageNumber = 1 }
            };
            RetrieveRecordChangeHistoryResponse response = service.Execute(request) as RetrieveRecordChangeHistoryResponse;
            var detailsToInclude = response.AuditDetailCollection.AuditDetails
                .Where(ad => ad is AttributeAuditDetail && ad.AuditRecord.Contains("createdon") && ((DateTime)ad.AuditRecord["createdon"]) > oldestUpdate)
                .OrderByDescending(ad => ((DateTime)ad.AuditRecord["createdon"]))
                .ToList();

            if (detailsToInclude.Any())
            {
                Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest retrieveEntityRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = workflowContext.PrimaryEntityName
                };
                Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse retrieveEntityResponse = service.Execute(retrieveEntityRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
                EntityMetadata metadata = retrieveEntityResponse.EntityMetadata;

                foreach (var detail in detailsToInclude.Select(d => d as AttributeAuditDetail).Where(d => d.NewValue != null && d.OldValue != null))
                {
                    DateTime dateToModify = (DateTime)detail.AuditRecord["createdon"];
                    if (dateToModify.Kind != DateTimeKind.Utc)
                    {
                        dateToModify = dateToModify.ToUniversalTime();
                    }


                    LocalTimeFromUtcTimeRequest timeZoneChangeRequest = new LocalTimeFromUtcTimeRequest() { UtcTime = dateToModify, TimeZoneCode = timeZone.MicrosoftIndex };
                    LocalTimeFromUtcTimeResponse timeZoneResponse = service.Execute(timeZoneChangeRequest) as LocalTimeFromUtcTimeResponse;
                    DateTime timeZoneSpecificDateTime = timeZoneResponse.LocalTime;

                    var details = detail.NewValue.Attributes.Keys.Union(detail.OldValue.Attributes.Keys)
                        .Distinct()
                        .Select(a =>
                        new {
                            AttributeName = a,
                            DisplayName = GetDisplayLabel(metadata, a)
                        })
                        .OrderBy(a => a.DisplayName);

                    foreach (var item in details)
                    {
                        DataRow newRow = table.NewRow();
                        newRow["User"] = GetDisplayValue(detail.AuditRecord, "userid");
                        newRow["Date"] = timeZoneSpecificDateTime.ToString("MM/dd/yyyy h:mm tt");
                        newRow["Attribute"] = item.DisplayName;
                        newRow["Old Value"] = GetDisplayValue(detail.OldValue, item.AttributeName);
                        newRow["New Value"] = GetDisplayValue(detail.NewValue, item.AttributeName);
                        table.Rows.Add(newRow);
                    }
                }
            }
            return table;
        }

        [Input("Max # of audit logs (up to 100)")]
        [RequiredArgument()]
        [Default("100")]
        public InArgument<int> MaxAuditLogs { get; set; }

        [RequiredArgument]
        [Input("Time Zone for Display")]
        [AttributeTarget("afk_workflowelementoption", "afk_extendedtimezones")]
        public InArgument<OptionSetValue> TimeZoneOption { get; set; }

        [RequiredArgument]
        [Input("Time Period Value")]
        public InArgument<int> Number { get; set; }

        [RequiredArgument]
        [Input("Time Period Unit")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeunits")]
        public InArgument<OptionSetValue> Units { get; set; }

        protected struct AttributeSummary
        {
            public string AttributeName { get; set; }
            public string DisplayName { get; set; }
            public string OldValue { get; set; }
            public string NewValue { get; set; }
        }
    }
}
