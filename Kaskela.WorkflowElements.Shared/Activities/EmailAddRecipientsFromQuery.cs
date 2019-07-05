using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class EmailAddRecipientsFromQuery : ContributingClasses.WorkflowQueryBase
    {
        [RequiredArgument]
        [Input("Email to Update")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> Email { get; set; }

        [RequiredArgument]
        [Input("Email Recipient Field")]
        [AttributeTarget("afk_workflowelementoption", "afk_emailrecipient")]
        public InArgument<OptionSetValue> EmailRecipientField { get; set; }

        [Output("Number of Recipients Added")]
        public OutArgument<int> NumberOfRecipientsAdded { get; set; }

        protected override void Execute(System.Activities.CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var service = this.RetrieveOrganizationService(context);

            string emailField = "to";
            switch (this.EmailRecipientField.Get(context).Value)
            {
                case 222540001:
                    emailField = "cc";
                    break;
                case 222540002:
                    emailField = "bcc";
                    break;
                default:
                    break;
            }

            Entity email = service.Retrieve("email", this.Email.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet(emailField, "statuscode"));
            if (email["statuscode"] == null || email.GetAttributeValue<OptionSetValue>("statuscode").Value != 1)
            {
                throw new ArgumentException("Email must be in Draft status.");
            }

            QueryResult result = ExecuteQueryForRecords(context);
            int recipientsAdded = result.RecordIds.Count();


            if (result.RecordIds.Any())
            {
                List<Entity> recipients = new List<Entity>();
                if (email[emailField] != null && ((Microsoft.Xrm.Sdk.EntityCollection)email[emailField]).Entities != null && ((Microsoft.Xrm.Sdk.EntityCollection)email[emailField]).Entities.Any())
                {
                    recipients.AddRange(((EntityCollection)email[emailField]).Entities.ToList());
                }

                foreach (Guid id in result.RecordIds)
                {
                    if (!recipients.Any(r => ((EntityReference)(r["partyid"])).Id == id && ((EntityReference)(r["partyid"])).LogicalName == result.EntityName))
                    {
                        Entity party = new Entity("activityparty");
                        party["partyid"] = new EntityReference(result.EntityName, id);
                        recipients.Add(party);
                        recipientsAdded++;
                    }
                }

                email[emailField] = recipients.ToArray();
            }
            service.Update(email);
            NumberOfRecipientsAdded.Set(context, recipientsAdded);
        }
    }
}