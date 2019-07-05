using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using Kaskela.WorkflowElements.Shared;
using Kaskela.WorkflowElements.Shared.ContributingClasses;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class DateGetDateComponents : WorkflowBase
    {
        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            DateTime utcDateTime = this.DateToEvaluate.Get(context);
            if (utcDateTime.Kind != DateTimeKind.Utc)
            {
                utcDateTime = utcDateTime.ToUniversalTime();
            }

            TimeZoneSummary timeZone = StaticMethods.CalculateTimeZoneToUse(this.TimeZoneOption.Get(context), workflowContext, service);
            LocalTimeFromUtcTimeRequest timeZoneChangeRequest = new LocalTimeFromUtcTimeRequest() { UtcTime = utcDateTime, TimeZoneCode = timeZone.MicrosoftIndex };
            LocalTimeFromUtcTimeResponse timeZoneResponse = service.Execute(timeZoneChangeRequest) as LocalTimeFromUtcTimeResponse;
            DateTime adjustedDateTime = timeZoneResponse.LocalTime;

            switch (adjustedDateTime.DayOfWeek)
            {
                case System.DayOfWeek.Sunday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540000));
                    break;
                case System.DayOfWeek.Monday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540001));
                    break;
                case System.DayOfWeek.Tuesday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540002));
                    break;
                case System.DayOfWeek.Wednesday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540003));
                    break;
                case System.DayOfWeek.Thursday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540004));
                    break;
                case System.DayOfWeek.Friday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540005));
                    break;
                case System.DayOfWeek.Saturday:
                    this.DayOfWeekPick.Set(context, new OptionSetValue(222540006));
                    break;
            }
            this.DayOfWeekName.Set(context, adjustedDateTime.DayOfWeek.ToString());
            this.DayOfMonth.Set(context, adjustedDateTime.Day);
            this.DayOfYear.Set(context, adjustedDateTime.DayOfYear);
            this.HourOfDay023.Set(context, adjustedDateTime.Hour);
            this.Minute.Set(context, adjustedDateTime.Minute);
            this.MonthOfYearInt.Set(context, adjustedDateTime.Month);
            this.MonthOfYearPick.Set(context, new OptionSetValue(222540000 + adjustedDateTime.Month - 1));
            this.MonthOfYearName.Set(context, System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(adjustedDateTime.Month));
            this.Year.Set(context, adjustedDateTime.Year);
        }

        [RequiredArgument]
        [Input("Date to Evaluate")]
        public InArgument<DateTime> DateToEvaluate { get; set; }

        [RequiredArgument]
        [Input("Time Zone to Use")]
        [AttributeTarget("afk_workflowelementoption", "afk_extendedtimezones")]
        public InArgument<OptionSetValue> TimeZoneOption { get; set; }

        [Output("Day of Week (Option Set)")]
        [AttributeTarget("afk_workflowelementoption", "afk_dayofweek")]
        public OutArgument<OptionSetValue> DayOfWeekPick { get; set; }

        [Output("Day of Week (Name)")]
        public OutArgument<string> DayOfWeekName { get; set; }

        [Output("Day of Month")]
        public OutArgument<int> DayOfMonth { get; set; }

        [Output("Day of Year")]
        public OutArgument<int> DayOfYear { get; set; }

        [Output("Month of Year (1 - 12)")]
        public OutArgument<int> MonthOfYearInt { get; set; }

        [Output("Month of Year (Option Set)")]
        [AttributeTarget("afk_workflowelementoption", "afk_monthofyear")]
        public OutArgument<OptionSetValue> MonthOfYearPick { get; set; }

        [Output("Month of Year (Name)")]
        public OutArgument<string> MonthOfYearName { get; set; }

        [Output("Year")]
        public OutArgument<int> Year { get; set; }
        
        [Output("Hour of the Day (0 - 23)")]
        public OutArgument<int> HourOfDay023 { get; set; }

        [Output("Minute")]
        public OutArgument<int> Minute { get; set; }

    }
}
