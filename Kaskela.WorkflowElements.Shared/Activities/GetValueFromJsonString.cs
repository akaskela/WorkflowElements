using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json.Linq;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class GetValueFromJson : ContributingClasses.WorkflowBase
    {
        [Input("Json String")]
        public InArgument<string> JsonString { get; set; }

        [Input("Json Path")]
        public InArgument<string> JsonPath { get; set; }

        [Output("Single Value - Result as Whole Number")]
        public OutArgument<int> Result_WholeNumber { get; set; }

        [Output("Single Value - Result as DateTime")]
        public OutArgument<DateTime> Result_DateTime { get; set; }

        [Output("Single Value - Result as Decimal")]
        public OutArgument<decimal> Result_Decimal { get; set; }

        [Output("Single Value - Result as Double")]
        public OutArgument<double> Result_Double { get; set; }

        [Output("Single Value - Result as Money")]
        public OutArgument<Money> Result_Money { get; set; }

        [Output("Single Value - Result as ID")]
        public OutArgument<string> Result_Guid { get; set; }

        [Output("Single Value - Result as Text")]
        public OutArgument<string> Result_Text { get; set; }


        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            if (workflowContext != null && workflowContext.InputParameters.Contains("Target") && workflowContext.InputParameters["Target"] is Entity && !string.IsNullOrEmpty(JsonString.Get<string>(context)) && !string.IsNullOrEmpty(JsonPath.Get<string>(context)))
            {
                var entity = (Entity)workflowContext.InputParameters["Target"];
                JObject jsonObject = JObject.Parse(JsonString.Get<string>(context));
                JToken token = jsonObject.SelectToken(JsonPath.Get<string>(context));
                if(token != null)
                {
                    string value = (string)token;
                    Result_Text.Set(context, value);

                    Guid id = new Guid();
                    if (Guid.TryParse(value, out id))
                    {
                        this.Result_Guid.Set(context, value);
                    }

                    DateTime dateTime = new DateTime();
                    if (DateTime.TryParse(value, out dateTime))
                    {
                        this.Result_DateTime.Set(context, dateTime);
                    }

                    decimal decimalValue = new decimal();
                    if (decimal.TryParse(value, out decimalValue))
                    {
                        this.Result_Decimal.Set(context, decimalValue);
                        try
                        {
                            this.Result_Money.Set(context, new Money(decimalValue));
                        }
                        catch { }
                    }

                    double doubleValue = new double();
                    if (double.TryParse(value, out doubleValue))
                    {
                        this.Result_Double.Set(context, doubleValue);
                    }

                    int intValue = 0;
                    if (int.TryParse(value, out intValue))
                    {
                        this.Result_WholeNumber.Set(context, intValue);
                    }
                }
            }
        }
    }
}
