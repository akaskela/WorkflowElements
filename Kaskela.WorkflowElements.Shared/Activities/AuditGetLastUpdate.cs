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
    public class AuditGetLastUpdate : ContributingClasses.AuditBase
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
            DataTable table = new DataTable() { TableName = workflowContext.PrimaryEntityName };
            table.Columns.AddRange(new DataColumn[] { new DataColumn("Attribute"), new DataColumn("Old Value"), new DataColumn("New Value") });

            RetrieveRecordChangeHistoryRequest request = new RetrieveRecordChangeHistoryRequest()
            {
                Target = new EntityReference(workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId),
                PagingInfo = new PagingInfo() { Count = 100, PageNumber = 1 }
            };
            RetrieveRecordChangeHistoryResponse response = service.Execute(request) as RetrieveRecordChangeHistoryResponse;
            
            if (response != null && response.AuditDetailCollection.Count > 0)
            {
                string onlyIfField = LastUpdateWithThisField.Get(context);
                AttributeAuditDetail detail = null;
                for (int i = 0; i < response.AuditDetailCollection.Count; i++)
                {
                    AttributeAuditDetail thisDetail = response.AuditDetailCollection[i] as AttributeAuditDetail;
                    if (thisDetail != null && (String.IsNullOrEmpty(onlyIfField) ||
                        (thisDetail.OldValue.Attributes.Keys.Contains(onlyIfField) ||
                        thisDetail.NewValue.Attributes.Keys.Contains(onlyIfField))))
                    {
                        detail = thisDetail;
                        break;
                    }
                }

                if (detail != null && detail.NewValue != null && detail.OldValue != null)
                {
                    Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest retrieveEntityRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest()
                    {
                        EntityFilters = EntityFilters.Attributes,
                        LogicalName = workflowContext.PrimaryEntityName
                    };
                    Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse retrieveEntityResponse = service.Execute(retrieveEntityRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
                    EntityMetadata metadata = retrieveEntityResponse.EntityMetadata;

                    var details = detail.NewValue.Attributes.Keys.Union(detail.OldValue.Attributes.Keys)
                        .Distinct()
                        .Select(a => new
                        {
                            AttributeName = a,
                            DisplayName = this.GetDisplayLabel(metadata, a),
                            OldValue = String.Empty,
                            NewValue = String.Empty
                        })
                        .OrderBy(a => a.DisplayName);

                    foreach (var item in details)
                    {
                        DataRow newRow = table.NewRow();
                        newRow["Attribute"] = item.DisplayName;
                        newRow["Old Value"] = GetDisplayValue(detail.OldValue, item.AttributeName);
                        newRow["New Value"] = GetDisplayValue(detail.NewValue, item.AttributeName);
                        table.Rows.Add(newRow);
                    }
                }
            }
            return table;
        }

        [Input("Last Update with This Field...")]
        public InArgument<string> LastUpdateWithThisField { get; set; }
        
        protected struct AttributeSummary
        {
            public string AttributeName { get; set; }
            public string DisplayName { get; set; }
            public string OldValue { get; set; }
            public string NewValue { get; set; }
        }
    }
}
