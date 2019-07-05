using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.CE.Activities
{
    public class SalesRemoveLineItems : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            if (this.SalesOrder.Get(context) == null && this.Quote.Get(context) == null &&
                this.Invoice.Get(context) == null && this.Opportunity.Get(context) == null)
            {
                throw new ArgumentNullException("You need to specify either an Opportunity, Quote, Order or Invoice");
            }
            
            string lineItemEntityName = String.Empty;
            string lineItemParentIdName = String.Empty;
            Guid lineItemParentId = Guid.Empty;
            
            if (this.Invoice.Get(context) != null)
            {
                lineItemEntityName = "invoicedetail";
                lineItemParentIdName = "invoiceid";
                lineItemParentId = this.Invoice.Get(context).Id;
            }
            else if (this.Opportunity.Get(context) != null)
            {
                lineItemEntityName = "opportunityproduct";
                lineItemParentIdName = "opportunityid";
                lineItemParentId = this.Opportunity.Get(context).Id;
            }
            else if (this.SalesOrder.Get(context) != null)
            {
                lineItemEntityName = "salesorderdetail";
                lineItemParentIdName = "salesorderid";
                lineItemParentId = this.SalesOrder.Get(context).Id;
            }
            else if (this.Quote.Get(context) != null)
            {
                lineItemEntityName = "quotedetail";
                lineItemParentIdName = "quoteid";
                lineItemParentId = this.Quote.Get(context).Id;
            }

            QueryExpression qe = new QueryExpression(lineItemEntityName);
            qe.Criteria.AddCondition(lineItemParentIdName, ConditionOperator.Equal, lineItemParentId);
            if (this.Product.Get(context) != null)
            {
                qe.Criteria.AddCondition("productid", ConditionOperator.Equal, this.Product.Get(context).Id);
            }

            var entities = service.RetrieveMultiple(qe).Entities;
            foreach (var lineItem in entities)
            {
                service.Delete(lineItem.LogicalName, lineItem.Id);
            }

            this.NumberRemoved.Set(context, entities.Count);
        }

        [Input("Invoice to Update")]
        [ReferenceTarget("invoice")]
        public InArgument<EntityReference> Invoice { get; set; }

        [Input("Opportunity to Update")]
        [ReferenceTarget("opportunity")]
        public InArgument<EntityReference> Opportunity { get; set; }

        [Input("Order to Update")]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> SalesOrder { get; set; }

        [Input("Quote to Update")]
        [ReferenceTarget("quote")]
        public InArgument<EntityReference> Quote { get; set; }
        
        [Input("Product to Remove (optional)")]
        [ReferenceTarget("product")]
        public InArgument<EntityReference> Product { get; set; }
        
        [Output("Number of Products Removed")]
        public OutArgument<int> NumberRemoved { get; set; }
    }
}