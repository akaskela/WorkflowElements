using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class DateBuildDateFromParts : WorkflowBase
    {
        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            int day = this.Day.Get(context);
            if (day <= 0 || day > 31)
            {
                throw new ArgumentOutOfRangeException("Day outside of valid range (1 - 31)");
            }
            int month = this.MonthOfYearInt.Get(context);
            if (month < 1 || month > 12)
            {
                month = this.MonthOfYearPick.Get(context).Value - 222540000 + 1;
            }

            int year = this.Year.Get(context);
            int hour = this.HourOfDay023.Get(context);
            int minute = this.Minute.Get(context);

            DateTime parsedDate = DateTime.MinValue;
            try
            {
                parsedDate = new DateTime(year, month, day, hour, minute, 0, DateTimeKind.Unspecified);
            }
            catch (Exception ex)
            {
                throw new ArgumentException("Error parsing date: " + ex.Message);
            }

            TimeZoneSummary timeZone = StaticMethods.CalculateTimeZoneToUse(this.TimeZoneOption.Get(context), workflowContext, service);
            UtcTimeFromLocalTimeRequest timeZoneChangeRequest = new UtcTimeFromLocalTimeRequest() { LocalTime = parsedDate, TimeZoneCode = timeZone.MicrosoftIndex };
            UtcTimeFromLocalTimeResponse timeZoneResponse = service.Execute(timeZoneChangeRequest) as UtcTimeFromLocalTimeResponse;
            DateTime adjustedDateTime = timeZoneResponse.UtcTime;

            this.ModifiedDate.Set(context, adjustedDateTime);
        }

        [Input("Day of Month")]
        public InArgument<int> Day { get; set; }

        [Input("Month (Option Set)")]
        [AttributeTarget("afk_workflowelementoption", "afk_monthofyear")]
        public InArgument<OptionSetValue> MonthOfYearPick { get; set; }

        [Input("Month of Year (1 - 12)")]
        public InArgument<int> MonthOfYearInt { get; set; }

        [Input("Year")]
        public InArgument<int> Year { get; set; }

        [Input("Hour of the Day (0 - 23)")]
        public InArgument<int> HourOfDay023 { get; set; }

        [Input("Minute")]
        public InArgument<int> Minute { get; set; }

        [Output("Date")]
        public OutArgument<DateTime> ModifiedDate { get; set; }

        [RequiredArgument]
        [Input("Time Zone to Use")]
        [AttributeTarget("afk_workflowelementoption", "afk_extendedtimezones")]
        public InArgument<OptionSetValue> TimeZoneOption { get; set; }
    }
}
