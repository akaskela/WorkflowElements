using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.CE.Activities
{
    public class SalesAddLineItem : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            this.Successful.Set(context, false);

            if (this.SalesOrder.Get(context) == null && this.Quote.Get(context) == null &&
                this.Invoice.Get(context) == null && this.Opportunity.Get(context) == null)
            {
                throw new ArgumentNullException("You need to specify either an Opportunity, Quote, Order or Invoice");
            }

            if ((this.Product.Get(context) == null || this.UOM.Get(context) == null) &&
                (this.ProductName.Get(context) == null || this.PricePerUnit.Get(context) == null))
            {
                throw new ArgumentNullException("You need to specify either a Product and Unit, or Product Name and Price per Unit");
            }

            Entity detail = null;
            if (this.Invoice.Get(context) != null)
            {
                detail = new Entity("invoicedetail");
                detail["invoiceid"] = new EntityReference("invoice", this.Invoice.Get(context).Id);
            }
            else if (this.Opportunity.Get(context) != null)
            {
                detail = new Entity("opportunityproduct");
                detail["opportunityid"] = new EntityReference("opportunity", this.Opportunity.Get(context).Id);
            }
            else if (this.SalesOrder.Get(context) != null)
            {
                detail = new Entity("salesorderdetail");
                detail["salesorderid"] = new EntityReference("salesorder", this.SalesOrder.Get(context).Id);
            }
            else if (this.Quote.Get(context) != null)
            {
                detail = new Entity("invoicedetail"); detail = new Entity("quotedetail");
                detail["quoteid"] = new EntityReference("quote", this.Quote.Get(context).Id);
            }

            if ((this.SalesOrder.Get(context) != null || this.Quote.Get(context) != null) &&
                (this.RequestDeliveryBy.Get(context) != DateTime.MinValue))
            {
                detail["requestdeliveryby"] = this.RequestDeliveryBy.Get(context);
            }
            
            detail["quantity"] = this.Quantity.Get(context);
            detail["description"] = this.Description.Get(context);

            if (this.Product.Get(context) != null && this.UOM.Get(context) != null)
            {
                detail["isproductoverridden"] = false;
                detail["productid"] = new EntityReference("product", this.Product.Get(context).Id);
                detail["uomid"] = new EntityReference("uom", this.UOM.Get(context).Id);
                if (this.PricePerUnit.Get(context) != null)
                {
                    detail["ispriceoverridden"] = true;
                    detail["priceperunit"] = this.PricePerUnit.Get(context);
                }
            }
            else
            {
                detail["isproductoverridden"] = true;
                detail["productdescription"] = this.ProductName.Get(context);
                detail["priceperunit"] = this.PricePerUnit.Get(context);
            }

            if (this.ManualDiscount.Get(context) != null)
            {
                detail["manualdiscountamount"] = this.ManualDiscount.Get(context);
            }
            if (this.Tax.Get(context) != null)
            {
                detail["tax"] = this.Tax.Get(context);
            }

            service.Create(detail);
            this.Successful.Set(context, true);
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
        
        [Input("Product to Add and")]
        [ReferenceTarget("product")]
        public InArgument<EntityReference> Product { get; set; }

        [Input("Unit of Measure")]
        [ReferenceTarget("uom")]
        public InArgument<EntityReference> UOM { get; set; }

        [Input("or Write-In Product Name and")]
        [ReferenceTarget("product")]
        public InArgument<string> ProductName { get; set; }

        [Input("Price per Unit")]
        public InArgument<Money> PricePerUnit { get; set; }

        [Input("Quantity")]
        [RequiredArgument()]
        public InArgument<decimal> Quantity { get; set; }

        [Input("Manual Discount")]
        public InArgument<Money> ManualDiscount { get; set; }

        [Input("Tax")]
        public InArgument<Money> Tax { get; set; }

        [Input("Description")]
        public InArgument<string> Description { get; set; }

        [Input("Request Delivery by (quote or order)")]
        public InArgument<DateTime> RequestDeliveryBy { get; set; }

        [Output("Product has been Added")]
        public OutArgument<bool> Successful { get; set; }

    }
}