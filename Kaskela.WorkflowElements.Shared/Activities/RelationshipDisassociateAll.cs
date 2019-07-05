using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Linq;
namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class RelationshipDisassociateAll : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        [RequiredArgument]
        [Input("Relationship Name")]
        public InArgument<string> RelationshipName { get; set; }
        
        [Output("# of Relationships Changes")]
        public OutArgument<int> NumberOfRelationshipChanges { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest metadataRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest() { LogicalName = workflowContext.PrimaryEntityName, EntityFilters = EntityFilters.Relationships };
            var metadataResponse = service.Execute(metadataRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
            var relationship = metadataResponse.EntityMetadata.ManyToManyRelationships.FirstOrDefault(m =>
                m.SchemaName.Equals(RelationshipName.Get(context), StringComparison.InvariantCultureIgnoreCase) ||
                m.IntersectEntityName.Equals(RelationshipName.Get(context), StringComparison.InvariantCultureIgnoreCase));
            if (relationship == null)
            {
                throw new Exception($"Entity '{workflowContext.PrimaryEntityName}' does not have relationship with schema name or relationship entity name '{RelationshipName.Get(context)}'");
            }

            QueryExpression qe = new QueryExpression(relationship.IntersectEntityName);
            qe.Criteria.AddCondition(relationship.Entity1IntersectAttribute, ConditionOperator.Equal, workflowContext.PrimaryEntityId);
            qe.ColumnSet = new ColumnSet(relationship.Entity2IntersectAttribute);
            var associated = service.RetrieveMultiple(qe).Entities.Select(e => new EntityReference(relationship.Entity2LogicalName, (Guid)e[relationship.Entity2IntersectAttribute])).ToList();

            service.Disassociate(workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId, new Relationship(relationship.SchemaName), new EntityReferenceCollection(associated));

            NumberOfRelationshipChanges.Set(context, associated.Count);
        }
    }
}