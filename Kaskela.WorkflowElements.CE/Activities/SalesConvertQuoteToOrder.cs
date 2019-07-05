using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.CE.Activities
{
    public class SalesConvertQuoteToOrder : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            Entity quote = service.Retrieve("quote", this.Quote.Get(context).Id, new ColumnSet("statuscode"));
            if (((OptionSetValue)(quote["statuscode"])).Value != 4)
            {
                throw new InvalidOperationException("The quote must be status 'Won' to convert to an order.");
            }
            ConvertQuoteToSalesOrderRequest convertQuoteRequest = new ConvertQuoteToSalesOrderRequest()
            {
                QuoteId = this.Quote.Get(context).Id,
                ColumnSet = new ColumnSet("salesorderid")
            };
            ConvertQuoteToSalesOrderResponse convertQuoteResponse = (ConvertQuoteToSalesOrderResponse)service.Execute(convertQuoteRequest);
            this.SalesOrder.Set(context, new EntityReference("salesorder", convertQuoteResponse.Entity.Id));
        }

        [Input("Quote to Convert")]
        [RequiredArgument()]
        [ReferenceTarget("quote")]
        public InArgument<EntityReference> Quote { get; set; }

        [Output("Sales Order")]
        [ReferenceTarget("salesorder")]
        public OutArgument<EntityReference> SalesOrder { get; set; }
    }
}