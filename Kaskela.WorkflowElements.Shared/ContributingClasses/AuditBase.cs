using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Kaskela.WorkflowElements.Shared.ContributingClasses
{
    public abstract class AuditBase : WorkflowBase
    {
        protected void ValidateColorCode(string value, string fieldName)
        {
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex("^#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})$");
            if (!regex.IsMatch(value))
            {
                throw new ArgumentException($"Hex code invalid for '{fieldName}' - value must be formatted as '#xxxxxx'");
            }
        }

        protected string GetDisplayValue(Entity e, string fieldName)
        {
            string returnValue = String.Empty;
            if (e.Contains(fieldName))
            {
                object itemToEvaluate = e[fieldName];
                if (itemToEvaluate is AliasedValue)
                {
                    itemToEvaluate = ((AliasedValue)e[fieldName]).Value;
                }

                if (itemToEvaluate is EntityReference)
                {
                    returnValue = ((EntityReference)itemToEvaluate).Name;
                }
                else if (itemToEvaluate is string)
                {
                    returnValue = itemToEvaluate.ToString();
                }
                else if (e.FormattedValues.Contains(fieldName))
                {
                    returnValue = e.FormattedValues[fieldName];
                }
                else
                {

                }
            }
            return returnValue;
        }
        protected string GetDisplayLabel(EntityMetadata metadata, string attributeName)
        {
            string returnValue = attributeName;

            AttributeMetadata attributeMetadata = metadata.Attributes.FirstOrDefault(a => a.LogicalName == attributeName);
            if (attributeMetadata != null)
            {
                Label displayLabel = attributeMetadata.DisplayName;
                if (displayLabel != null && displayLabel.UserLocalizedLabel != null && !String.IsNullOrWhiteSpace(displayLabel.UserLocalizedLabel.Label))
                {
                    returnValue = displayLabel.UserLocalizedLabel.Label;
                }
            }

            return returnValue;
        }

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
                    string fontWeight = this.Header_BoldFont.Get(context) ? "bold" : "normal";
                    sb.Append($"<td style=\"border: 1px solid {borderColor}; padding: 6px; background-color: {this.Header_BackgroundColor.Get(context)}; color: {this.Header_FontColor.Get(context)}; font-weight: {fontWeight}\">");
                    sb.Append(column.ColumnName);
                    sb.Append("</td>");
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
                    sb.Append($"<td style=\"border: 1px solid {borderColor}; padding: 6px; \">");
                    sb.Append(table.Rows[rowNumber][i].ToString());
                    sb.Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</table>");

            this.HtmlAuditResults.Set(context, sb.ToString());
        }
        protected virtual void BuildResultsAsCsv(DataTable table, CodeActivityContext context)
        {
            List<string> rows = new List<string>();
            List<string> currentRow = new List<string>();
            if (this.IncludeHeader.Get(context))
            {
                foreach (DataColumn column in table.Columns)
                {
                    currentRow.Add(column.ColumnName);
                }
                rows.Add(String.Join(",", currentRow));
            }

            for (int rowNumber = 0; rowNumber < table.Rows.Count; rowNumber++)
            {
                currentRow = new List<string>();
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    currentRow.Add(table.Rows[rowNumber][i].ToString());
                }
                rows.Add(String.Join(",", currentRow));
            }
            this.CsvAuditResults.Set(context, String.Join(Environment.NewLine, rows));
        }

        [Input("Include Header for Table?")]
        [Default("True")]
        public InArgument<Boolean> IncludeHeader { get; set; }

        [RequiredArgument]
        [Input("Table - Border Color")]
        [Default("#000000")]
        public InArgument<string> Table_BorderColor { get; set; }

        [RequiredArgument]
        [Input("Header - Background Color")]
        [Default("#6495ED")]
        public InArgument<string> Header_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Header - Font Color")]
        [Default("#FFFFFF")]
        public InArgument<string> Header_FontColor { get; set; }

        [RequiredArgument]
        [Input("Header - Make Font Bold")]
        [Default("True")]
        public InArgument<Boolean> Header_BoldFont { get; set; }

        [RequiredArgument]
        [Input("Row - Background Color")]
        [Default("#FFFFFF")]
        public InArgument<string> Row_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Row - Font Color")]
        [Default("#000000")]
        public InArgument<string> Row_FontColor { get; set; }

        [RequiredArgument]
        [Input("Alternating Row - Background Color")]
        [Default("#F0F8FF")]
        public InArgument<string> AlternatingRow_BackgroundColor { get; set; }

        [RequiredArgument]
        [Input("Alternating Row - Font Color")]
        [Default("#000000")]
        public InArgument<string> AlternatingRow_FontColor { get; set; }


        [Output("Audit History (HTML)")]
        public OutArgument<string> HtmlAuditResults { get; set; }

        [Output("Audit History (CSV)")]
        public OutArgument<string> CsvAuditResults { get; set; }

        [Output("Number of Fields in Audit Result")]
        public OutArgument<int> NumberOfFields { get; set; }
    }
}
