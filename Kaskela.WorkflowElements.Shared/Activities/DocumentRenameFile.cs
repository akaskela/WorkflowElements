using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class DocumentRenameFile : ContributingClasses.RetrieveActivityBase
    {
        protected override void Execute(CodeActivityContext context)
        {
            this.Successful.Set(context, false);

            Entity note = base.RetrieveAnnotationEntity(context, new Microsoft.Xrm.Sdk.Query.ColumnSet("filename", "mimetype", "documentbody")).FirstOrDefault();
            if (note != null)
            {
                IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

                note["filename"] = NewFilename.Get(context);
                service.Update(note);
                this.Successful.Set(context, true);
            }
        }


        [Input("New filename (including extension)")]
        [RequiredArgument()]
        public InArgument<string> NewFilename { get; set; }

        [Output("Document Has Been Updated")]
        public OutArgument<bool> Successful { get; set; }
    }
}