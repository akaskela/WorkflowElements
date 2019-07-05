using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class QueryRunWorkflowOnResults : ContributingClasses.WorkflowQueryBase
    {
        [RequiredArgument]
        [Input("Workflow to Run")]
        [ReferenceTarget("workflow")]
        public InArgument<EntityReference> Workflow { get; set; }
        
        [Output("Number of Workflows Started")]
        public OutArgument<int> NumberOfWorkflowsStarted { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            QueryResult result = ExecuteQueryForRecords(context);
            Entity workflow = service.Retrieve("workflow", this.Workflow.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("primaryentity"));

            if (workflow.GetAttributeValue<string>("primaryentity").ToLower() != result.EntityName)
            {
                throw new ArgumentException($"Workflow entity ({workflow.GetAttributeValue<string>("primaryentity")} does not match query entity ({result.EntityName})");
            }

            int numberStarted = 0;
            foreach (Guid id in result.RecordIds)
            {
                try
                {
                    ExecuteWorkflowRequest request = new ExecuteWorkflowRequest() { EntityId = id, WorkflowId = workflow.Id };
                    ExecuteWorkflowResponse response = service.Execute(request) as ExecuteWorkflowResponse;
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error initiating workflow on record with ID = {id}; {ex.Message}");
                }
                numberStarted++;
            }
            this.NumberOfWorkflowsStarted.Set(context, numberStarted);
        }
    }
}