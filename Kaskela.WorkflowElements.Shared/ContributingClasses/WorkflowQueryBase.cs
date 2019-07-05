using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;

namespace Kaskela.WorkflowElements.Shared.ContributingClasses
{
    public abstract class WorkflowQueryBase : WorkflowBase
    {
        protected string MetadataHeaderFormat = "{0}_Metadata_{1}";

        [Input("Pick a System View to Use")]
        [ReferenceTarget("savedquery")]
        public InArgument<EntityReference> SavedQuery { get; set; }

        [Input("or pick a Personal View to Use")]
        [ReferenceTarget("userquery")]
        public InArgument<EntityReference> UserQuery { get; set; }

        [Input("or enter FetchXML to Use")]
        public InArgument<string> FetchXml { get; set; }

        protected DataTable ExecuteQuery(CodeActivityContext context)
        {
            if (this.SavedQuery.Get(context) == null && this.UserQuery.Get(context) == null && String.IsNullOrEmpty(this.FetchXml.Get(context)))
            {
                throw new ArgumentNullException("You need to either a pick System View, a Personal View, or specify FetchXML for the query");
            }

            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);

            string layoutXml = String.Empty;
            string fetchXml = this.FetchXml.Get(context);
            if (this.UserQuery.Get(context) != null)
            {
                Entity queryEntity = service.Retrieve("userquery", this.UserQuery.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("fetchxml", "layoutxml"));
                fetchXml = queryEntity.Contains("fetchxml") ? queryEntity["fetchxml"].ToString() : String.Empty;
                layoutXml = queryEntity.Contains("layoutxml") ? queryEntity["layoutxml"].ToString() : String.Empty;
            }
            else if (this.SavedQuery.Get(context) != null)
            {
                Entity queryEntity = service.Retrieve("savedquery", this.SavedQuery.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("fetchxml", "layoutxml"));
                fetchXml = queryEntity.Contains("fetchxml") ? queryEntity["fetchxml"].ToString() : String.Empty;
                layoutXml = queryEntity.Contains("layoutxml") ? queryEntity["layoutxml"].ToString() : String.Empty;
            }

            this.LimitQueryToCurrentRecord(service, ref fetchXml, workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId);

            List<ColumnInformation> tableColumns = this.BuildTableInformation(service, fetchXml, layoutXml);
            
            EntityCollection ec = null;
            try
            {
                FetchExpression fe = new FetchExpression(fetchXml);
                ec = service.RetrieveMultiple(fe);
            }
            catch (System.ServiceModel.FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                throw new ArgumentException("There was an error executing the query. Message: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            DataTable table = new DataTable() { TableName = ec.EntityName };
            table.Columns.AddRange(tableColumns.Select(c => new DataColumn(c.Header)).ToArray());
            foreach (var entity in ec.Entities)
            {
                DataRow newRow = table.NewRow();
                foreach (ColumnInformation c in tableColumns.Where(c => !c.Header.Contains("_Metadata_")))
                {
                    if (entity.Contains(c.QueryResultAlias))
                    {
                        object itemToEvaluate = entity[c.QueryResultAlias];
                        if (itemToEvaluate is AliasedValue)
                        {
                            itemToEvaluate = ((AliasedValue)entity[c.QueryResultAlias]).Value;
                        }

                        if (itemToEvaluate is EntityReference)
                        {
                            newRow[c.Header] = ((EntityReference)itemToEvaluate).Name;
                            newRow[String.Format(MetadataHeaderFormat, c.Header, "RawValue")] = ((EntityReference)itemToEvaluate).Id.ToString();
                        }
                        else if (itemToEvaluate is string)
                        {
                            newRow[c.Header] = itemToEvaluate.ToString();
                            newRow[String.Format(MetadataHeaderFormat, c.Header, "RawValue")] = itemToEvaluate.ToString();
                        }
                        else if (entity.FormattedValues.Contains(c.QueryResultAlias))
                        {
                            newRow[c.Header] = entity.FormattedValues[c.QueryResultAlias];

                            if (itemToEvaluate is Money)
                            {
                                newRow[String.Format(MetadataHeaderFormat, c.Header, "RawValue")] = ((Money)itemToEvaluate).Value;
                            }
                            else
                            {
                                newRow[String.Format(MetadataHeaderFormat, c.Header, "RawValue")] = entity.FormattedValues[c.QueryResultAlias];
                            }

                        }
                        else if (itemToEvaluate is Guid)
                        {
                            newRow[c.Header] = (Guid)itemToEvaluate;
                        }
                    }
                    else
                    {
                        newRow[c.Header] = String.Empty;
                        newRow[String.Format(MetadataHeaderFormat, c.Header, "RawValue")] = String.Empty;
                    }
                }
                table.Rows.Add(newRow);
            }
            return table;
        }

        protected QueryResult ExecuteQueryForRecords(CodeActivityContext context)
        {
            QueryResult returnValue = new QueryResult() { RecordIds = new List<Guid>() };
            if (this.SavedQuery.Get(context) == null && this.UserQuery.Get(context) == null && String.IsNullOrEmpty(this.FetchXml.Get(context)))
            {
                throw new ArgumentNullException("You need to either a pick System View, a Personal View, or specify FetchXML for the query");
            }

            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);


            string layoutXml = String.Empty;
            string fetchXml = this.FetchXml.Get(context);
            if (this.UserQuery.Get(context) != null)
            {
                Entity queryEntity = service.Retrieve("userquery", this.UserQuery.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("fetchxml", "layoutxml"));
                fetchXml = queryEntity.Contains("fetchxml") ? queryEntity["fetchxml"].ToString() : String.Empty;
                layoutXml = queryEntity.Contains("layoutxml") ? queryEntity["layoutxml"].ToString() : String.Empty;
            }
            else if (this.SavedQuery.Get(context) != null)
            {
                Entity queryEntity = service.Retrieve("savedquery", this.SavedQuery.Get(context).Id, new Microsoft.Xrm.Sdk.Query.ColumnSet("fetchxml", "layoutxml"));
                fetchXml = queryEntity.Contains("fetchxml") ? queryEntity["fetchxml"].ToString() : String.Empty;
                layoutXml = queryEntity.Contains("layoutxml") ? queryEntity["layoutxml"].ToString() : String.Empty;
            }
            this.LimitQueryToCurrentRecord(service, ref fetchXml, workflowContext.PrimaryEntityName, workflowContext.PrimaryEntityId);
            EntityCollection ec = null;
            try
            {
                FetchExpression fe = new FetchExpression(fetchXml);
                ec = service.RetrieveMultiple(fe);
            }
            catch (System.ServiceModel.FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> ex)
            {
                throw new ArgumentException("There was an error executing the query. Message: " + ex.Message);
            }
            catch (Exception ex)
            {
                throw ex;
            }

            returnValue.EntityName = ec.EntityName;

            RetrieveEntityRequest entityMetadataRequest = new RetrieveEntityRequest() { LogicalName = returnValue.EntityName };
            var response = service.Execute(entityMetadataRequest) as RetrieveEntityResponse;
            if (response == null || response.EntityMetadata == null)
            {
                throw new InvalidOperationException($"Entity metadata not found for entity {returnValue.EntityName}");
            }
            returnValue.Metadata = response.EntityMetadata;

            string primaryKeyField = returnValue.Metadata.PrimaryIdAttribute;

            returnValue.RecordIds.AddRange(ec.Entities.Select(e => e.Id).ToList());

            return returnValue;
        }

        protected List<ColumnInformation> BuildTableInformation(IOrganizationService service, string fetchXml, string layoutXml)
        {
            List<QueryAttribute> fetchXmlColumns = new List<QueryAttribute>();
            XmlDocument doc = new XmlDocument();
            doc.LoadXml(fetchXml);
            XmlNode entityNode = doc["fetch"]["entity"];

            string primaryEntity = entityNode.Attributes["name"].Value;
            this.RetrieveQueryAttributeRecursively(entityNode, fetchXmlColumns);

            List<QueryAttribute> layoutXmlColumns = new List<QueryAttribute>();
            if (!String.IsNullOrEmpty(layoutXml))
            {
                doc = new XmlDocument();
                doc.LoadXml(layoutXml);
                foreach (XmlNode cell in doc["grid"]["row"].ChildNodes)
                {
                    string columnPresentationValue = cell.Attributes["name"].Value;
                    if (columnPresentationValue.Contains("."))
                    {
                        layoutXmlColumns.Add(new QueryAttribute()
                        {
                            EntityAlias = columnPresentationValue.Split('.')[0],
                            Attribute = columnPresentationValue.Split('.')[1]
                        });
                    }
                    else
                    {
                        layoutXmlColumns.Add(new QueryAttribute() { Attribute = columnPresentationValue, EntityName = primaryEntity });
                    }
                }
                foreach (QueryAttribute attribute in layoutXmlColumns.Where(qa => !String.IsNullOrWhiteSpace(qa.EntityAlias)))
                {
                    QueryAttribute attributeFromFetch = fetchXmlColumns.FirstOrDefault(f => f.EntityAlias == attribute.EntityAlias);
                    if (attributeFromFetch != null)
                    {
                        attribute.EntityName = attributeFromFetch.EntityName;
                    }
                }
            }

            List<string> distintEntityNames = layoutXmlColumns.Select(lo => lo.EntityName).Union(fetchXmlColumns.Select(f => f.EntityName)).Distinct().ToList();
            foreach (string entityName in distintEntityNames)
            {
                Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest request = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest()
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = entityName
                };
                Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse response = service.Execute(request) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
                EntityMetadata metadata = response.EntityMetadata;
                foreach (QueryAttribute queryAttribute in layoutXmlColumns.Union(fetchXmlColumns).Where(c => c.EntityName == entityName && !String.IsNullOrWhiteSpace(c.Attribute)))
                {
                    if (!String.IsNullOrWhiteSpace(queryAttribute.AttribueAlias))
                    {
                        queryAttribute.AttributeLabel = queryAttribute.AttribueAlias;
                    }
                    else
                    {
                        AttributeMetadata attributeMetadata = metadata.Attributes.FirstOrDefault(a => a.LogicalName == queryAttribute.Attribute);
                        Label displayLabel = attributeMetadata.DisplayName;
                        if (displayLabel != null && displayLabel.UserLocalizedLabel != null)
                        {
                            queryAttribute.AttributeLabel = displayLabel.UserLocalizedLabel.Label;
                        }
                    }
                }
            }

            List<ColumnInformation> returnColumns = new List<ColumnInformation>();
            List<QueryAttribute> toReturn = (layoutXmlColumns.Count > 0) ? layoutXmlColumns : fetchXmlColumns;
            foreach (QueryAttribute qa in toReturn)
            {
                ColumnInformation ci = new ColumnInformation();
                ci.Header = qa.AttributeLabel;
                if (!String.IsNullOrEmpty(qa.EntityAlias))
                {
                    ci.QueryResultAlias = String.Format("{0}.{1}", qa.EntityAlias, qa.Attribute);
                }
                else if (!String.IsNullOrEmpty(qa.AttribueAlias))
                {
                    ci.QueryResultAlias = qa.AttribueAlias;
                }
                else
                {
                    ci.QueryResultAlias = qa.Attribute;
                }
                returnColumns.Add(ci);
                returnColumns.Add(new ColumnInformation() { Header = String.Format(MetadataHeaderFormat, ci.Header, "RawValue") });
            }

            return returnColumns;
        }

        protected void RetrieveQueryAttributeRecursively(XmlNode node, List<QueryAttribute> attribues)
        {
            string entityAlias = String.Empty;
            string entityName = String.Empty;
            if (node.Name == "entity")
            {
                entityName = node.Attributes["name"].Value;
                entityAlias = String.Empty;
            }
            else if (node.Name == "link-entity")
            {
                entityName = node.Attributes["name"].Value;
                entityAlias = node.Attributes["alias"] != null ? node.Attributes["alias"].Value : entityName;
            }
            foreach (XmlNode child in node.ChildNodes)
            {
                if (child.Name == "attribute")
                {
                    attribues.Add(new QueryAttribute()
                    {
                        AttribueAlias = child.Attributes["alias"] != null ? child.Attributes["alias"].Value : String.Empty,
                        EntityName = entityName,
                        Attribute = child.Attributes["name"].Value,
                        EntityAlias = entityAlias
                    });
                }
                else if (child.Name == "link-entity")
                {
                    this.RetrieveQueryAttributeRecursively(child, attribues);
                }
            }
        }

        protected void LimitQueryToCurrentRecord(IOrganizationService service, ref string fetchXml, string entityName, Guid entityId)
        {
            Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest metadataRequest = new Microsoft.Xrm.Sdk.Messages.RetrieveEntityRequest() { LogicalName = entityName };
            Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse metadataResponse = service.Execute(metadataRequest) as Microsoft.Xrm.Sdk.Messages.RetrieveEntityResponse;
            if (metadataResponse == null)
            {
                return;
            }

            XmlDocument fetchXmlDoc = new XmlDocument();
            fetchXmlDoc.LoadXml(fetchXml);
            this.UpdateConditionRecursively(fetchXmlDoc, fetchXmlDoc.FirstChild, entityId, entityName, metadataResponse.EntityMetadata.PrimaryIdAttribute);
            fetchXml = fetchXmlDoc.OuterXml;
        }

        protected void UpdateConditionRecursively(XmlDocument xmlDoc, XmlNode currentNode, Guid entityId, string entityName, string primaryKeyField)
        {
            if (currentNode.Name == "condition")
            {
                if (currentNode.Attributes["attribute"] != null && currentNode.Attributes["attribute"].Value == primaryKeyField &&
                    currentNode.Attributes["operator"] != null && currentNode.Attributes["operator"].Value == "not-null" &&
                    currentNode.ParentNode != null && currentNode.ParentNode.Name == "filter" &&
                    currentNode.ParentNode.ParentNode != null && currentNode.ParentNode.ParentNode.Name == "link-entity" &&
                    currentNode.ParentNode.ParentNode.Attributes["name"] != null && currentNode.ParentNode.ParentNode.Attributes["name"].Value == entityName)
                {
                    XmlAttribute keyAttribute = xmlDoc.CreateAttribute("value");
                    keyAttribute.Value = entityId.ToString();
                    currentNode.Attributes["operator"].Value = "eq";
                    currentNode.Attributes.Append(keyAttribute);
                }
            }
            foreach (XmlNode child in currentNode.ChildNodes)
            {
                this.UpdateConditionRecursively(xmlDoc, child, entityId, entityName, primaryKeyField);
            }
        }
        
        public struct QueryResult
        {
            public string EntityName { get; set; }
            public List<Guid> RecordIds { get; set; }
            public EntityMetadata Metadata { get; set; }
        }
    }
}
