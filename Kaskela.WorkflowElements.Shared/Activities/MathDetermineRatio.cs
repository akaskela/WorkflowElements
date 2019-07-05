using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class MathDetermineRatio : WorkflowBase
    {
        [RequiredArgument]
        [Input("Value 1")]
        public InArgument<Decimal> DecimalValue1 { get; set; }

        [RequiredArgument]
        [Input("Value 2")]
        public InArgument<Decimal> DecimalValue2 { get; set; }

        [Input("Number of Significant Digits")]
        [AttributeTarget("afk_workflowelementoption", "afk_zerototen")]
        [Default("222540099")]
        public InArgument<OptionSetValue> SignificantDigits { get; set; }

        [Input("Number of Decimal Places")]
        [AttributeTarget("afk_workflowelementoption", "afk_zerototen")]
        [Default("222540099")]
        public InArgument<OptionSetValue> DecimalPlaces { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            decimal value1 = this.DecimalValue1.Get(context);
            decimal value2 = this.DecimalValue2.Get(context);
            
            if (value1 <= 0 || value2 <= 0)
            {
                this.RatioEvaluated.Set(context, false);
                //this.RatioExpression.Set(context, "N/A");
            }
            else
            {
                // Evaluate VALUE:1 ratio
                decimal ratioExpressionPart1 = value1 / value2;
                decimal ratioExpressionPart2 = value2 / value1;

                if (this.DecimalPlaces.Get(context) != null && this.DecimalPlaces.Get(context).Value != 222540099)
                {
                    int decimalPlaces = this.DecimalPlaces.Get(context).Value - 222540000;
                    ratioExpressionPart2 = new NumericManipulator().RoundToDecimalPlaces(ratioExpressionPart2, decimalPlaces);
                    ratioExpressionPart1 = new NumericManipulator().RoundToDecimalPlaces(ratioExpressionPart1, decimalPlaces);
                }

                if (this.SignificantDigits.Get(context) != null && this.SignificantDigits.Get(context).Value != 222540099)
                {
                    int significantDigits = this.SignificantDigits.Get(context).Value - 222540000;
                    ratioExpressionPart2 = new NumericManipulator().RoundToSignificantDigits(ratioExpressionPart2, significantDigits);
                    ratioExpressionPart1 = new NumericManipulator().RoundToSignificantDigits(ratioExpressionPart2, significantDigits);
                }                
                this.RatioEvaluated.Set(context, true);


                this.RatioValueToOne.Set(context, ratioExpressionPart1);
                this.RatioExpressionValueToOne.Set(context, ($"{ratioExpressionPart1}:1"));

                this.RatioOneToValue.Set(context, ratioExpressionPart2);
                this.RatioExpressionOneToValue.Set(context, ($"1:{ratioExpressionPart2}"));
            }
        }

        [Output("Ratio Value (value1:1)")]
        public OutArgument<Decimal> RatioValueToOne { get; set; }
        
        [Output("Ratio Value (1:value2)")]
        public OutArgument<Decimal> RatioOneToValue { get; set; }

        [Output("Ratio Expression (value1:1)")]
        public OutArgument<string> RatioExpressionValueToOne { get; set; }

        [Output("Ratio Expression (1:value2)")]
        public OutArgument<string> RatioExpressionOneToValue { get; set; }

        [Output("Ratio Evaluated")]
        public OutArgument<bool> RatioEvaluated { get; set; }
    }
}
