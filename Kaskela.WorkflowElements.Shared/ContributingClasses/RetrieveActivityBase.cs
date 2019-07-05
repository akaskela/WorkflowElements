using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Xml;

namespace Kaskela.WorkflowElements.Shared.ContributingClasses
{
    public abstract class RetrieveActivityBase : WorkflowBase
    {
        [Input("Filename must contain... (optional)")]
        public InArgument<string> FileName { get; set; }

        [Input("Only as far back as {value}")]
        [RequiredArgument()]
        [Default("5")]
        public InArgument<int> TimeSpanValue { get; set; }

        [Input("Only as far back as {units}")]
        [RequiredArgument()]
        [AttributeTarget("afk_workflowelementoption", "afk_timespanoption")]
        public InArgument<OptionSetValue> TimeSpanOption { get; set; }


        protected List<Entity> RetrieveAnnotationEntity(CodeActivityContext context, ColumnSet noteColumns, int maxRecords = 1)
        {
            double miunutesOld = 0;
            List<Entity> returnValue = new List<Entity>();
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

            int? objectTypeCode = this.RetrieveEntityObjectTypeCode(workflowContext, service);
            if (objectTypeCode == null)
            {
                throw new ArgumentException($"Objecttypecode not found in metadata for entity {workflowContext.PrimaryEntityName}");
            }

            ExecuteFetchResponse fetchResponse = null;
            ExecuteFetchRequest request = new ExecuteFetchRequest();
            try
            {
                if (String.IsNullOrWhiteSpace(this.FileName.Get(context)))
                {
                    request.FetchXml =
                        $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"" page=""1"" count=""{maxRecords}"">
                              <entity name=""annotation"">
                                <attribute name=""annotationid"" />
                                <attribute name=""createdon"" />
                                <filter type=""and"">
                                  <condition attribute=""isdocument"" operator=""eq"" value=""1"" />
                                  <condition attribute=""objectid"" operator=""eq"" value=""{workflowContext.PrimaryEntityId}"" />
                                  <condition attribute=""objecttypecode"" operator=""eq"" value=""{objectTypeCode.Value}"" />
                                </filter>
                                <order attribute=""createdon"" descending=""true"" />
                              </entity>
                            </fetch>";
                }
                else
                {
                    request.FetchXml =
                            $@"<fetch version=""1.0"" output-format=""xml-platform"" mapping=""logical"" distinct=""false"" page=""1"" count=""{maxRecords}"">
                              <entity name=""annotation"">
                                <attribute name=""annotationid"" />
                                <attribute name=""createdon"" />
                                <filter type=""and"">
                                  <condition attribute=""filename"" operator=""like"" value=""%{this.FileName.Get(context)}%"" />
                                  <condition attribute=""isdocument"" operator=""eq"" value=""1"" />
                                  <condition attribute=""objectid"" operator=""eq"" value=""{workflowContext.PrimaryEntityId}"" />
                                  <condition attribute=""objecttypecode"" operator=""eq"" value=""{objectTypeCode.Value}"" />
                                </filter>
                                <order attribute=""createdon"" descending=""true"" />
                              </entity>
                            </fetch>";
                }

                fetchResponse = service.Execute(request) as ExecuteFetchResponse;

                XmlDocument queryResults = new XmlDocument();
                queryResults.LoadXml(fetchResponse.FetchXmlResult);
                int days = 0;
                int minutes = 0;
                for (int i = 0; i < queryResults["resultset"].ChildNodes.Count; i++)
                {
                    if (queryResults["resultset"].ChildNodes[i]["createdon"] != null && !String.IsNullOrWhiteSpace(queryResults["resultset"].ChildNodes[i]["createdon"].InnerText))
                    {
                        DateTime createdon = DateTime.Parse(queryResults["resultset"].ChildNodes[i]["createdon"].InnerText);
                        if (createdon.Kind == DateTimeKind.Local)
                        {
                            createdon = createdon.ToUniversalTime();
                        }
                        TimeSpan difference = DateTime.Now.ToUniversalTime() - createdon;
                        miunutesOld = difference.TotalMinutes;
                        switch (this.TimeSpanOption.Get(context).Value)
                        {
                            case 222540000:
                                minutes = this.TimeSpanValue.Get(context);
                                break;
                            case 222540001:
                                minutes = this.TimeSpanValue.Get(context) * 60;
                                break;
                            case 222540002:
                                days = this.TimeSpanValue.Get(context);
                                break;
                            case 222540003:
                                days = this.TimeSpanValue.Get(context) * 7;
                                break;
                            case 222540004:
                                days = this.TimeSpanValue.Get(context) * 365;
                                break;
                        }
                        TimeSpan allowedDifference = new TimeSpan(days, 0, minutes, 0);
                        if (difference <= allowedDifference)
                        {
                            returnValue.Add(service.Retrieve("annotation", Guid.Parse(queryResults["resultset"].ChildNodes[i]["annotationid"].InnerText), noteColumns));
                        }
                    }

                    if (returnValue.Count >= maxRecords)
                    {
                        break;
                    }
                }
            }
            catch (System.ServiceModel.FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                throw new ArgumentException("There was an error executing the FetchXML. Message: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return returnValue;
        }

        protected int? RetrieveEntityObjectTypeCode(IWorkflowContext workflowContext, IOrganizationService service)
        {
            int? returnValue = null;

            Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest metadataRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest() { LogicalName = workflowContext.PrimaryEntityName };
            Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse metadataResponse = service.Execute(metadataRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
            if (metadataResponse != null)
            {
                returnValue = metadataResponse.EntityMetadata.ObjectTypeCode;
            }

            return returnValue;
        }
    }
}
