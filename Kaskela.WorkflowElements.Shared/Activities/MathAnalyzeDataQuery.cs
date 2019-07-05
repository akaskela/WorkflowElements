using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class MathAnalyzeDataQuery : ContributingClasses.WorkflowQueryBase
    {

        [Input("Number of Significant Digits")]
        [AttributeTarget("afk_workflowelementoption", "afk_zerototen")]
        [Default("222540099")]
        public InArgument<OptionSetValue> SignificantDigits { get; set; }

        [Input("Number of Decimal Places")]
        [AttributeTarget("afk_workflowelementoption", "afk_zerototen")]
        [Default("222540099")]
        public InArgument<OptionSetValue> DecimalPlaces { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            DataTable table = this.ExecuteQuery(context);
            if (table.Columns == null || table.Columns.Count == 0)
            {
                throw new ArgumentException("Query result must have at least one column");
            }


            List<decimal> inputValues = new List<decimal>();
            for (int rowNumber = 0; rowNumber < table.Rows.Count; rowNumber++)
            {
                decimal parsedValue = 0;
                if (decimal.TryParse(table.Rows[rowNumber][1].ToString(), out parsedValue))
                {
                    inputValues.Add(parsedValue);
                }
            }

            this.NumberOfItems.Set(context, inputValues.Count);
            if (inputValues.Any())
            {
                this.Mean.Set(context, this.RoundValue(context, inputValues.Average()));
                this.Median.Set(context, this.RoundValue(context, this.CalculateMedian(inputValues)));
                this.Highest.Set(context, this.RoundValue(context, inputValues.Max()));
                this.Lowest.Set(context, this.RoundValue(context, inputValues.Min()));
                this.StandardDeviation.Set(context, this.RoundValue(context, CalculateStandardDeviation(inputValues)));
            }
        }
        protected decimal CalculateMedian(List<decimal> input)
        {
            List<decimal> values = input.OrderBy(n => n).ToList();
            decimal returnValue = 0;
            int numberCount = values.Count();
            int halfIndex = values.Count() / 2;
            if ((numberCount % 2) == 0)
            {
                decimal element1 = values.ElementAt(halfIndex);
                decimal element2 = values.ElementAt(halfIndex - 1);
                returnValue = (element1 + element2) / 2;
            }
            else
            {
                returnValue = values.ElementAt(halfIndex);
            }
            return returnValue;
        }
        protected decimal CalculateStandardDeviation(List<decimal> input)
        {
            decimal meanToUse = input.Average(i => i);

            if (input.Count == 1)
            {
                return 0;
            }
            else
            {
                decimal sumOfDifferences = 0;
                foreach (decimal value in input)
                {
                    decimal difference = (value - meanToUse);
                    sumOfDifferences += difference * difference;
                }
                decimal variance = sumOfDifferences / (input.Count - 1);
                return (decimal)Math.Sqrt((double)variance);
            }
        }
        protected decimal RoundValue(CodeActivityContext context, decimal input)
        {
            decimal roundedValue = input;
            if (this.DecimalPlaces.Get(context) != null && this.DecimalPlaces.Get(context).Value != 222540099)
            {
                int decimalPlaces = this.DecimalPlaces.Get(context).Value - 222540000;
                roundedValue = new NumericManipulator().RoundToDecimalPlaces(roundedValue, decimalPlaces);
            }

            if (this.SignificantDigits.Get(context) != null && this.SignificantDigits.Get(context).Value != 222540099)
            {
                int significantDigits = this.SignificantDigits.Get(context).Value - 222540000;
                roundedValue = new NumericManipulator().RoundToSignificantDigits(roundedValue, significantDigits);
            }
            return roundedValue;
        }

        [Output("Mean")]
        public OutArgument<decimal> Mean { get; set; }

        [Output("Median")]
        public OutArgument<decimal> Median { get; set; }

        [Output("Highest Number")]
        public OutArgument<decimal> Highest { get; set; }

        [Output("Lowest Number")]
        public OutArgument<decimal> Lowest { get; set; }

        [Output("Standard Deviation")]
        public OutArgument<decimal> StandardDeviation { get; set; }
        
        [Output("Number of items in Data Set")]
        public OutArgument<int> NumberOfItems { get; set; }
    }
}
