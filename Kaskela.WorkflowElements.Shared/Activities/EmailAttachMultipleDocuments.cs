using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class EmailAttachMultipleDocuments : ContributingClasses.RetrieveActivityBase
    {
        [Input("Email to Attach Files")]
        [RequiredArgument()]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

        [Input("Max # of attachments (up to 20)")]
        [RequiredArgument()]
        [Default("20")]
        public InArgument<int> MaxAttachments { get; set; }

        [Input("Filename can't contain... (optional)")]
        public InArgument<string> FileNameCantHave { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            this.AttachmentCount.Set(context, 0);

            int max = this.MaxAttachments.Get(context);
            if (max > 20 || max < 1)
            {
                max = 20;
            }

            List<Entity> notes = base.RetrieveAnnotationEntity(context, new Microsoft.Xrm.Sdk.Query.ColumnSet("filename", "mimetype", "documentbody"), max);
            var potentialAttachments = notes.Select(note => new
            {
                FileName = note["filename"].ToString(),
                MimeType = note["mimetype"],
                DocumentBody = note["documentbody"]
            }).ToList();

            if (!String.IsNullOrWhiteSpace(this.FileNameCantHave.Get(context)))
            {
                List<string> portions = this.FileNameCantHave.Get(context).ToLower().Split('|').ToList();
                portions.ForEach(p => potentialAttachments.RemoveAll(pa => pa.FileName.ToLower().Contains(p)));
            }

            if (potentialAttachments.Any())
            {
                IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

                foreach (var a in potentialAttachments)
                {
                    Entity attachment = new Entity("activitymimeattachment");
                    attachment["objectid"] = new EntityReference("email", this.Email.Get(context).Id);
                    attachment["objecttypecode"] = "email";
                    attachment["filename"] = a.FileName.ToString();
                    attachment["mimetype"] = a.MimeType;
                    attachment["body"] = a.DocumentBody;
                    service.Create(attachment);
                }
                this.AttachmentCount.Set(context, potentialAttachments.Count());
            }
        }



        [Output("# of Documents Attached")]
        public OutArgument<int> AttachmentCount { get; set; }
    }
}