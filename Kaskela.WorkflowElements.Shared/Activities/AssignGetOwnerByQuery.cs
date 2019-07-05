using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class AssignGetOwnerByQuery : ContributingClasses.WorkflowQueryBase
    {
        [Output("Record Reassigned")]
        public OutArgument<bool> RecordReassigned { get; set; }

        [Output("New Owner (User)")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> NewOwner_User { get; set; }

        [Output("New Owner (Team)")]
        [ReferenceTarget("team")]
        public OutArgument<EntityReference> NewOwner_Team { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            QueryResult result = ExecuteQueryForRecords(context);
            if (!result.EntityName.Equals("systemuser", StringComparison.InvariantCultureIgnoreCase) && 
                !result.EntityName.Equals("team", StringComparison.InvariantCultureIgnoreCase))
            {
                throw new ArgumentException("Query must return a User or Team record");
            }

            // Ensure the record has an owner field
            RetrieveEntityRequest request = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest()
            {
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity,
                LogicalName = workflowContext.PrimaryEntityName
            };
            RetrieveEntityResponse metadataResponse = service.Execute(request) as RetrieveEntityResponse;
            if (metadataResponse.EntityMetadata.OwnershipType == null || metadataResponse.EntityMetadata.OwnershipType.Value != OwnershipTypes.UserOwned)
            {
                throw new ArgumentException("This activity is only available for User owned records");
            }

            if (!result.RecordIds.Any())
            {
                return;
            }

            var workflowRecord = new EntityReference(workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId);
            var assignee = new EntityReference(result.EntityName, result.RecordIds.FirstOrDefault());
            AssignRequest assignRequest = new AssignRequest()
            {
                Assignee = assignee,
                Target = workflowRecord
            };

            try
            {
                AssignResponse response = service.Execute(assignRequest) as AssignResponse;
                if (assignee.LogicalName.Equals("team"))
                {
                    this.NewOwner_Team.Set(context, assignee);
                }
                else
                {
                    this.NewOwner_User.Set(context, assignee);
                }
                this.RecordReassigned.Set(context, true);
            }
            catch (Exception ex)
            {
                throw new Exception($"There was an error reassigning the record: {ex.Message}");
            }
        }
    }
}