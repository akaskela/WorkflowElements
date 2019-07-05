using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class DateConvertToCustomText : WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            List<OptionSetValue> values = new List<OptionSetValue>()  {
                this.Part1.Get<OptionSetValue>(context),
                this.Part2.Get<OptionSetValue>(context),
                this.Part3.Get<OptionSetValue>(context),
                this.Part4.Get<OptionSetValue>(context),
                this.Part5.Get<OptionSetValue>(context),
                this.Part6.Get<OptionSetValue>(context),
                this.Part7.Get<OptionSetValue>(context),
                this.Part8.Get<OptionSetValue>(context),
                this.Part9.Get<OptionSetValue>(context),
                this.Part10.Get<OptionSetValue>(context),
                this.Part11.Get<OptionSetValue>(context),
                this.Part12.Get<OptionSetValue>(context),
                this.Part13.Get<OptionSetValue>(context),
                this.Part14.Get<OptionSetValue>(context),
                this.Part15.Get<OptionSetValue>(context),
                this.Part16.Get<OptionSetValue>(context),
                this.Part17.Get<OptionSetValue>(context),
                this.Part18.Get<OptionSetValue>(context),
                this.Part19.Get<OptionSetValue>(context),
                this.Part20.Get<OptionSetValue>(context) };
            values.RemoveAll(osv => osv == null || osv.Value == 222540025);

            TimeZoneSummary timeZone = StaticMethods.CalculateTimeZoneToUse(this.TimeZoneOption.Get(context), workflowContext, service);
            DateTime dateToModify = this.DateToModify.Get(context);

            if (dateToModify.Kind != DateTimeKind.Utc)
            {
                dateToModify = dateToModify.ToUniversalTime();
            }

            
            LocalTimeFromUtcTimeRequest timeZoneChangeRequest = new LocalTimeFromUtcTimeRequest() { UtcTime = dateToModify, TimeZoneCode = timeZone.MicrosoftIndex };
            LocalTimeFromUtcTimeResponse timeZoneResponse = service.Execute(timeZoneChangeRequest) as LocalTimeFromUtcTimeResponse;
            DateTime timeZoneSpecificDateTime = timeZoneResponse.LocalTime;
            
            StringBuilder sb = new StringBuilder();
            
            foreach (OptionSetValue osv in values)
            {
                try
                {
                    switch (osv.Value)
                    {
                        case 222540000:
                            sb.Append(timeZoneSpecificDateTime.ToString("%h"));
                            break;
                        case 222540001:
                            sb.Append(timeZoneSpecificDateTime.ToString("hh"));
                            break;
                        case 222540002:
                            sb.Append(timeZoneSpecificDateTime.ToString("%H"));
                            break;
                        case 222540003:
                            sb.Append(timeZoneSpecificDateTime.ToString("HH"));
                            break;
                        case 222540004:
                            sb.Append(timeZoneSpecificDateTime.ToString("%m"));
                            break;
                        case 222540005:
                            sb.Append(timeZoneSpecificDateTime.ToString("mm"));
                            break;
                        case 222540006:
                            sb.Append(timeZoneSpecificDateTime.ToString("%d"));
                            break;
                        case 222540007:
                            sb.Append(timeZoneSpecificDateTime.ToString("dd"));
                            break;
                        case 222540008:
                            sb.Append(timeZoneSpecificDateTime.ToString("ddd"));
                            break;
                        case 222540009:
                            sb.Append(timeZoneSpecificDateTime.ToString("dddd"));
                            break;
                        case 222540010:
                            sb.Append(timeZoneSpecificDateTime.ToString("%M"));
                            break;
                        case 222540011:
                            sb.Append(timeZoneSpecificDateTime.ToString("MM"));
                            break;
                        case 222540012:
                            sb.Append(timeZoneSpecificDateTime.ToString("MMM"));
                            break;
                        case 222540013:
                            sb.Append(timeZoneSpecificDateTime.ToString("MMMM"));
                            break;
                        case 222540014:
                            sb.Append(timeZoneSpecificDateTime.ToString("%y"));
                            break;
                        case 222540015:
                            sb.Append(timeZoneSpecificDateTime.ToString("yy"));
                            break;
                        case 222540016:
                            sb.Append(timeZoneSpecificDateTime.ToString("yyyy"));
                            break;
                        case 222540017:
                            sb.Append(timeZoneSpecificDateTime.ToString("%t"));
                            break;
                        case 222540018:
                            sb.Append(timeZoneSpecificDateTime.ToString("tt"));
                            break;
                        case 222540019:
                            sb.Append(" ");
                            break;
                        case 222540020:
                            sb.Append(",");
                            break;
                        case 222540021:
                            sb.Append(".");
                            break;
                        case 222540022:
                            sb.Append(":");
                            break;
                        case 222540023:
                            sb.Append("/");
                            break;
                        case 222540024:
                            sb.Append(@"\");
                            break;
                        case 222540026:
                            sb.Append("-");
                            break;
                        case 222540027:
                            sb.Append(timeZone.Id);
                            break;
                        case 222540028:
                            sb.Append(timeZone.FullName);
                            break;
                        default:
                            break;
                    }
                }
                catch
                {
                    throw new Exception(osv.Value.ToString());
                }
            }
            
            FormattedDate.Set(context, sb.ToString());
        }

        [RequiredArgument]
        [Input("Date to Modify")]
        public InArgument<DateTime> DateToModify { get; set; }

        [RequiredArgument]
        [Input("Time Zone for Display")]
        [AttributeTarget("afk_workflowelementoption", "afk_extendedtimezones")]
        //[Default("222540000")]
        public InArgument<OptionSetValue> TimeZoneOption { get; set; }

        [RequiredArgument]
        [Input("First part 1")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        //[Default("222540006")]
        public InArgument<OptionSetValue> Part1 { get; set; }
        
        
        [Input("...then part 2")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part2 { get; set; }

        [Input("...then part 3")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part3 { get; set; }

        [Input("...then part 4")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part4 { get; set; }

        [Input("...then part 5")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part5 { get; set; }

        [Input("...then part 6")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part6 { get; set; }

        [Input("...then part 7")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part7 { get; set; }

        [Input("...then part 8")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part8 { get; set; }

        [Input("...then part 9")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part9 { get; set; }

        [Input("...then part 10")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part10 { get; set; }

        [Input("...then part 11")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part11 { get; set; }

        [Input("...then part 12")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part12 { get; set; }

        [Input("...then part 13")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part13 { get; set; }
        
        [Input("...then part 14")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part14 { get; set; }

        [Input("...then part 15")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part15 { get; set; }

        [Input("...then part 16")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part16 { get; set; }

        [Input("...then part 17")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part17 { get; set; }

        [Input("...then part 18")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part18 { get; set; }

        [Input("...then part 19")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part19 { get; set; }

        [Input("...then part 20")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeformats")]
        public InArgument<OptionSetValue> Part20 { get; set; }

        [Output("Formatted Date")]
        public OutArgument<string> FormattedDate { get; set; }
    }
}
