using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class QueryGetResults : ContributingClasses.WorkflowQueryBase
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

            DataTable table = this.ExecuteQuery(context);

            this.NumberOfResults.Set(context, table.Rows.Count);
            this.BuildResultsAsCsv(table, context);
            this.BuildResultsAsHtml(table, context, service);
            this.BuildListResults(table, context);
            this.SetSingleValueResults(table, context);
        }

        #region Validate Input
        protected void ValidateColorCode(string value, string fieldName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$");
            if (!regex.IsMatch(value))
            {
                throw new ArgumentException($"Hex code invalid for '{fieldName}' - value must be formatted as '#xxxxxx'");
            }
        }
        #endregion

        #region Set results
        protected virtual void BuildResultsAsHtml(DataTable table, CodeActivityContext context, IOrganizationService service)
        {
            string borderColor = this.Table_BorderColor.Get(context);
            StringBuilder sb = new StringBuilder();
            sb.Append($"<table style=\"border: 1px solid {borderColor}; border-collapse:collapse;\">");
            if (this.IncludeHeader.Get(context))
            {
                sb.Append("<tr>");
                foreach (DataColumn column in table.Columns)
                {
                    if (!column.ColumnName.Contains("_Metadata_"))
                    {
                        string fontWeight = this.Header_BoldFont.Get(context) ? "bold" : "normal";
                        sb.Append($"<td style=\"border: 1px solid {borderColor}; padding: 6px; background-color: {this.Header_BackgroundColor.Get(context)}; color: {this.Header_FontColor.Get(context)}; font-weight: {fontWeight}\">");
                        sb.Append(column.ColumnName);
                        sb.Append("</td>");
                    }
                }
                sb.Append("</tr>");
            }
            for (int rowNumber = 0; rowNumber < table.Rows.Count; rowNumber++)
            {
                string rowBackGroundColor = rowNumber % 2 == 0 ? this.Row_BackgroundColor.Get(context) : AlternatingRow_BackgroundColor.Get(context);
                string rowTextColor = rowNumber % 2 == 0 ? this.Row_FontColor.Get(context) : this.AlternatingRow_FontColor.Get(context);
                sb.Append($"<tr style=\"background-color: {rowBackGroundColor}; color: {rowTextColor}\">");
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (!table.Columns[i].ColumnName.Contains("_Metadata_"))
                    {
                        sb.Append($"<td style=\"border: 1px solid {borderColor}; padding: 6px; \">");
                        sb.Append(table.Rows[rowNumber][i].ToString());
                        sb.Append("</td>");
                    }
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");

            this.HtmlQueryResults.Set(context, sb.ToString());
        }
        protected virtual void BuildResultsAsCsv(DataTable table, CodeActivityContext context)
        {
            List<string> rows = new List<string>();
            List<string> currentRow = new List<string>();
            if (this.IncludeHeader.Get(context))
            {
                foreach (DataColumn column in table.Columns)
                {
                    if (!column.ColumnName.Contains("_Metadata_"))
                    {
                        currentRow.Add(column.ColumnName);
                    }
                }
                rows.Add(String.Join(",", currentRow));
            }

            for (int rowNumber = 0; rowNumber < table.Rows.Count; rowNumber++)
            {
                currentRow = new List<string>();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    if (!table.Columns[i].ColumnName.Contains("_Metadata_"))
                    {
                        currentRow.Add(table.Rows[rowNumber][i].ToString());
                    }
                }
                rows.Add(String.Join(",", currentRow));
            }
            this.CsvQueryResults.Set(context, String.Join(Environment.NewLine, rows));
        }
        protected virtual void BuildListResults(DataTable table, CodeActivityContext context)
        {
            List<string> listItems = new List<string>();
            for (int rowNumber = 0; rowNumber < table.Rows.Count; rowNumber++)
            {
                listItems.Add(table.Rows[rowNumber][0].ToString());
            }

            if (!this.IncludeEmptyItem.Get(context))
            {
                listItems.RemoveAll(i => String.IsNullOrEmpty(i));
            }

            this.NumberOfResults.Set(context, listItems.Count);
            this.ListResults.Set(context, String.Join(this.ListSeparator.Get(context), listItems));
        }
        protected virtual void SetSingleValueResults(DataTable table, CodeActivityContext context)
        {
            string innerValue = String.Empty;
            if (table.Rows.Count > 0)
            {
                this.QueryResult_Text.Set(context, table.Rows[0][0].ToString());
                innerValue = table.Rows[0][1].ToString();
            }

            Guid id = new Guid();
            if (Guid.TryParse(innerValue, out id))
            {
                this.QueryResult_Guid.Set(context, innerValue);
            }

            DateTime dateTime = new DateTime();
            if (DateTime.TryParse(innerValue, out dateTime))
            {
                this.QueryResult_DateTime.Set(context, dateTime);
            }

            decimal decimalValue = new decimal();
            if (decimal.TryParse(innerValue, out decimalValue))
            {
                this.QueryResult_Decimal.Set(context, decimalValue);
                try
                {
                    this.QueryResult_Money.Set(context, new Money(decimalValue));
                }
                catch { }
            }

            double doubleValue = new double();
            if (double.TryParse(innerValue, out doubleValue))
            {
                this.QueryResult_Double.Set(context, doubleValue);
            }

            int intValue = 0;
            if (int.TryParse(innerValue, out intValue))
            {
                this.QueryResult_WholeNumber.Set(context, intValue);
            }
        }
        #endregion
        
        #region Input Arguments

        [Input("Table - Include Header")]
        [Default("True")]
        public InArgument<Boolean> IncludeHeader { get; set; }

        [RequiredArgument]
        [Input("Table - Border Color")]
        [Default("#000000")]
        public InArgument<string> Table_BorderColor { get; set; }

        [RequiredArgument]
        [Input("Table - Header Background Color")]
        [Default("#6495ED")]
        public InArgument<string> Header_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Table - Header Font Color")]
        [Default("#FFFFFF")]
        public InArgument<string> Header_FontColor { get; set; }

        [RequiredArgument]
        [Input("Table - Header Bold Font?")]
        [Default("True")]
        public InArgument<Boolean> Header_BoldFont { get; set; }

        [RequiredArgument]
        [Input("Table - Row Background Color")]
        [Default("#FFFFFF")]
        public InArgument<string> Row_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Table - Row Font Color")]
        [Default("#000000")]
        public InArgument<string> Row_FontColor { get; set; }

        [RequiredArgument]
        [Input("Table - Alternating Row Background Color")]
        [Default("#F0F8FF")]
        public InArgument<string> AlternatingRow_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Table - Alternating Row Font Color")]
        [Default("#000000")]
        public InArgument<string> AlternatingRow_FontColor { get; set; }

        [Input("List - Item separator")]
        [Default(", ")]
        public InArgument<string> ListSeparator { get; set; }

        [Input("List - Include empty items?")]
        [Default("False")]
        public InArgument<Boolean> IncludeEmptyItem { get; set; }

        #endregion

        #region Output Arguments
        [Output("# of Results")]
        public OutArgument<int> NumberOfResults { get; set; }

        [Output("Table - Query Results (HTML)")]
        public OutArgument<string> HtmlQueryResults { get; set; }

        [Output("Table - Query Results (CSV)")]
        public OutArgument<string> CsvQueryResults { get; set; }
        
        [Output("List - Query Results as a List")]
        public OutArgument<string> ListResults { get; set; }

        [Output("Single Value - Result as Whole Number")]
        public OutArgument<int> QueryResult_WholeNumber { get; set; }

        [Output("Single Value - Result as DateTime")]
        public OutArgument<DateTime> QueryResult_DateTime { get; set; }

        [Output("Single Value - Result as Decimal")]
        public OutArgument<decimal> QueryResult_Decimal { get; set; }

        [Output("Single Value - Result as Double")]
        public OutArgument<double> QueryResult_Double { get; set; }

        [Output("Single Value - Result as Money")]
        public OutArgument<Money> QueryResult_Money { get; set; }

        [Output("Single Value - Result as ID")]
        public OutArgument<string> QueryResult_Guid { get; set; }

        [Output("Single Value - Result as Text")]
        public OutArgument<string> QueryResult_Text { get; set; }
        #endregion
    }
}
