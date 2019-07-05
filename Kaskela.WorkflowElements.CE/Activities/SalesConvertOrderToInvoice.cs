using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kaskela.WorkflowElements.CE.Activities
{
    public class SalesConvertOrderToInvoice : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            ConvertSalesOrderToInvoiceRequest convertOrderRequest = new ConvertSalesOrderToInvoiceRequest()
            {
                SalesOrderId = this.SalesOrder.Get(context).Id,
                ColumnSet = new ColumnSet("invoiceid")
            };
            ConvertSalesOrderToInvoiceResponse convertOrderResponse = (ConvertSalesOrderToInvoiceResponse)service.Execute(convertOrderRequest);
            this.Invoice.Set(context, new EntityReference("invoice", convertOrderResponse.Entity.Id));
        }

        [Input("Order to Convert")]
        [RequiredArgument()]
        [ReferenceTarget("salesorder")]
        public InArgument<EntityReference> SalesOrder { get; set; }

        [Output("Invoice")]
        [ReferenceTarget("invoice")]
        public OutArgument<EntityReference> Invoice { get; set; }
    }
}