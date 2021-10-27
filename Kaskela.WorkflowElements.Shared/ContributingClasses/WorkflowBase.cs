using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;

namespace Kaskela.WorkflowElements.Shared.ContributingClasses
{
    public abstract class WorkflowBase : CodeActivity
    {
        [RequiredArgument]
        [Input("Execution User")]
        [AttributeTarget("afk_workflowelementoption", "afk_useroption")]
        public InArgument<OptionSetValue> ExecutionUser { get; set; }

        protected IOrganizationService RetrieveOrganizationService(CodeActivityContext context)
        {
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            if (this.ExecutionUser.Get(context) == null)
            {
                throw new Exception("Execution User is required for this action.");
            }

            IOrganizationService returnValue = null;
            if (this.ExecutionUser.Get(context).Value == 222540000)
            {
                returnValue = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);
            }
            else if (this.ExecutionUser.Get(context).Value == 222540001)
            {
                returnValue = serviceFactory.CreateOrganizationService(workflowContext.UserId);
            }

            return returnValue;
        }
    }
}
