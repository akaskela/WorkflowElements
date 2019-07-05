using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class EmailSendSavedEmail : Kaskela.WorkflowElements.Shared.ContributingClasses.WorkflowBase
    {
        [RequiredArgument]
        [Input("Email to Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }
        
        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            Entity email = service.Retrieve("email", this.Email.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("statuscode"));
            if (email["statuscode"] == null || email.GetAttributeValue<OptionSetValue>("statuscode").Value != 1)
            {
                throw new ArgumentException("Email must be in Draft status.");
            }

            SendEmailRequest sendEmailrequest = new SendEmailRequest
            {
                EmailId = email.Id,
                TrackingToken = "",
                IssueSend = true
            };

            try
            {
                SendEmailResponse sendEmailresponse = (SendEmailResponse)service.Execute(sendEmailrequest);
            }
            catch(Exception ex)
            {
                throw new Exception($"Exception on SendEmailResponse - {ex.Message}");
            }
        }
    }
}