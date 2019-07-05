using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class WorkflowGetMetadata : ContributingClasses.WorkflowBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            this.WorkflowUser.Set(context, StaticMethods.RetrieveWorkflowUser(workflowContext));
            this.WorkflowDepth.Set(context, workflowContext.Depth);

            this.BusinessUnit.Set(context, new EntityReference("businessunit", workflowContext.BusinessUnitId));

            this.IsExecutingOffline.Set(context, workflowContext.IsExecutingOffline);

            if (workflowContext.IsolationMode == 1)
            {
                this.IsolationMode_Display.Set(context, "None");
                this.IsolationMode.Set(context, new OptionSetValue(222540001));
            }
            else if (workflowContext.IsolationMode == 2)
            {
                this.IsolationMode_Display.Set(context, "Sandbox");
                this.IsolationMode.Set(context, new OptionSetValue(222540000));
            }

            if (workflowContext.Mode == 0)
            {
                this.SynchronousMode_Display.Set(context, "Synchronous");
                this.SynchronousMode.Set(context, new OptionSetValue(222540000));
            }
            else if (workflowContext.Mode == 1)
            {
                this.SynchronousMode_Display.Set(context, "Asynchronous");
                this.SynchronousMode.Set(context, new OptionSetValue(222540001));
            }

            if (workflowContext.WorkflowCategory == 0)
            {
                this.WorkflowCategory_Display.Set(context, "Workflow");
                this.WorkflowCategory.Set(context, new OptionSetValue(222540000));
            }
            else if (workflowContext.IsolationMode == 1)
            {
                this.WorkflowCategory_Display.Set(context, "Dialog");
                this.WorkflowCategory.Set(context, new OptionSetValue(222540001));
            }

            this.OrganizationName.Set(context, workflowContext.OrganizationName);
            this.OrganizationID.Set(context, workflowContext.OrganizationId.ToString());

            this.SetWorkflowRecordInfo(context, workflowContext, service);
        }

        protected void SetWorkflowRecordInfo(CodeActivityContext context, IWorkflowContext workflowContext, IOrganizationService service)
        {
            // Record info
            this.WorkflowRecord_LookupID.Set(context, workflowContext.PrimaryEntityId.ToString());
            this.WorkflowRecord_LookupEntityName.Set(context, workflowContext.PrimaryEntityName);

            Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest metadataRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest() { LogicalName = workflowContext.PrimaryEntityName };
            Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse metadataResponse = service.Execute(metadataRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
            if (metadataResponse != null)
            {
                this.WorkflowRecord_LookupObjectTypeCode.Set(context, metadataResponse.EntityMetadata.ObjectTypeCode);
                if (metadataResponse.EntityMetadata.DisplayName != null && metadataResponse.EntityMetadata.DisplayName.UserLocalizedLabel != null)
                {
                    this.WorkflowRecord_LookupEntityDisplayName.Set(context, metadataResponse.EntityMetadata.DisplayName.UserLocalizedLabel.Label);
                }
            }
        }

        [Output("Business Unit")]
        [ReferenceTarget("businessunit")]
        public OutArgument<EntityReference> BusinessUnit { get; set; }

        [Output("Organization Name")]
        public OutArgument<string> OrganizationName { get; set; }

        [Output("Organization ID")]
        public OutArgument<string> OrganizationID { get; set; }

        [Output("Is Executing Offline")]
        public OutArgument<bool> IsExecutingOffline { get; set; }

        [Output("Isolation Mode as Text")]
        public OutArgument<string> IsolationMode_Display { get; set; }

        [Output("Isolation Mode as Option Set")]
        [AttributeTarget("afk_workflowelementoption", "afk_isolationmode")]
        public OutArgument<OptionSetValue> IsolationMode { get; set; }

        [Output("Synchronous Mode as Text")]
        [AttributeTarget("afk_workflowelementoption", "afk_synchronousmode")]
        public OutArgument<string> SynchronousMode_Display { get; set; }

        [Output("Synchronous Mode as Option Set")]
        [AttributeTarget("afk_workflowelementoption", "afk_synchronousmode")]
        public OutArgument<OptionSetValue> SynchronousMode { get; set; }

        [Output("Workflow Category as Text")]
        [AttributeTarget("afk_workflowelementoption", "afk_workflowcategory")]
        public OutArgument<string> WorkflowCategory_Display { get; set; }

        [Output("Workflow Category as Option Set")]
        [AttributeTarget("afk_workflowelementoption", "afk_workflowcategory")]
        public OutArgument<OptionSetValue> WorkflowCategory { get; set; }

        [Output("Workflow User")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> WorkflowUser { get; set; }

        [Output("Workflow Depth")]
        public OutArgument<int> WorkflowDepth { get; set; }

        [Output("Workflow Record - Entity Schema Name")]
        public OutArgument<string> WorkflowRecord_LookupEntityName { get; set; }

        [Output("Workflow Record - Entity Display Name")]
        public OutArgument<string> WorkflowRecord_LookupEntityDisplayName { get; set; }

        [Output("Workflow Record - Object Type Code")]
        public OutArgument<int> WorkflowRecord_LookupObjectTypeCode { get; set; }

        [Output("Workflow Record - Record ID")]
        public OutArgument<string> WorkflowRecord_LookupID { get; set; }
    }
}
