using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Linq;

namespace Kaskela.WorkflowElements.Shared
{
    public class StaticMethods
    {
        public static EntityReference RetrieveWorkflowRecordOwner(IWorkflowContext workflowContext, IOrganizationService service)
        {
            EntityReference recordOwner = null;
            RetrieveEntityRequest request = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest()
            {
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes,
                LogicalName = workflowContext.PrimaryEntityName
            };

            RetrieveEntityResponse metadataResponse = service.Execute(request) as RetrieveEntityResponse;
            LookupAttributeMetadata ownerAttribute = metadataResponse.EntityMetadata.Attributes.FirstOrDefault(att => att.AttributeType != null && (int)att.AttributeType.Value == 9) as LookupAttributeMetadata;
            if (ownerAttribute != null)
            {
                Entity entity = workflowContext.PostEntityImages.Values.FirstOrDefault();
                if (entity == null)
                {
                    entity = workflowContext.PreEntityImages.Values.FirstOrDefault();
                    if (entity == null)
                    {
                        entity = service.Retrieve(workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId, new ColumnSet(ownerAttribute.LogicalName));
                    }
                }

                if (entity != null && entity.Contains(ownerAttribute.LogicalName))
                {
                    recordOwner = entity[ownerAttribute.LogicalName] as EntityReference;
                }
            }

            return recordOwner;
        }

        public static EntityReference RetrieveWorkflowUser(IWorkflowContext workflowContext)
        {
            return new EntityReference("systemuser", workflowContext.InitiatingUserId);
        }
        
        public static TimeZoneSummary RetrieveTimeZoneForUser(EntityReference reference, IOrganizationService service)
        {
            TimeZoneSummary summary = null;
            Entity userSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.Equal, reference.Id)
                        }
                    }
                }).Entities.FirstOrDefault();
            if (userSettings != null && userSettings.Contains("timezonecode"))
            {
                summary = TimeZoneSummary.RetrieveTimeZoneByIndex(int.Parse(userSettings["timezonecode"].ToString()));
            }
            return summary;
        }
        
        public static TimeZoneSummary CalculateTimeZoneToUse(OptionSetValue timeZoneOption, IWorkflowContext workflowContext, IOrganizationService service)
        {
            TimeZoneSummary timeZoneSummary = null;
            if (timeZoneOption.Value == 222540001)
            {
                EntityReference owner = StaticMethods.RetrieveWorkflowRecordOwner(workflowContext, service);
                if (owner == null)
                {
                    throw new ArgumentException("Owner not found - \"The Owner of the Record\" is not valid for this entity");
                }
                timeZoneSummary = StaticMethods.RetrieveTimeZoneForUser(owner, service);
            }
            else if (timeZoneOption.Value == 222540000)
            {
                EntityReference owner = StaticMethods.RetrieveWorkflowUser(workflowContext);
                timeZoneSummary = StaticMethods.RetrieveTimeZoneForUser(owner, service);
            }
            else
            {
                timeZoneSummary = TimeZoneSummary.RetrieveTimeZoneByOptionSetValue(timeZoneOption.Value);
            }

            if (timeZoneSummary == null)
            {
                throw new ArgumentException("Time zone specified is invalid");
            }

            return timeZoneSummary;
        }
    }
}
