using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading.Tasks;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class EntitySpecification
    {
        public string EntitySchemaName { get; set; }
        public Guid EntityId { get; set; }
    }
    public class WebHook : ContributingClasses.WorkflowBase
    {
        [Input("RequestHeaders")]
        public InArgument<string> RequestHeaders { get; set; }

        [RequiredArgument]
        [Input("Request Url")]
        public InArgument<string> RequestUrl { get; set; }

        [Input("Request Body")]
        public InArgument<string> RequestBody { get; set; }

        public const int GET = 100000000;
        public const int POST = 100000001;
        public const int PUT = 100000002;
        public const int PATCH = 100000003;
        public const int DELETE = 100000004;

        [RequiredArgument]
        [Input("Request Method")]
        [AttributeTarget("afk_workflowelementoption", "new_httprequestmethod")]
        public InArgument<OptionSetValue> RequestMethod { get; set; }

        [Output("Response Headers")]
        public OutArgument<string> ResponseHeaders { get; set; }

        [Output("Response Body")]
        public OutArgument<string> ResponseBody { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            string body = string.Empty;
            var workflowContext = context.GetExtension<IWorkflowContext>();
            if (workflowContext != null && workflowContext.InputParameters.Contains("Target") && workflowContext.InputParameters["Target"] is Entity && string.IsNullOrEmpty(RequestBody.Get<string>(context)))
            {
                var entity = (Entity)workflowContext.InputParameters["Target"];
                body = Stringify(new EntitySpecification() { EntityId = entity.Id });
            }
            else
            {
                body = RequestBody.Get<string>(context);
            }
            using (var client = new HttpClient())
            {
                var content = new StringContent(body, Encoding.UTF8, "application/json");

                System.Threading.Tasks.Task<HttpResponseMessage> response = null;
                if (this.RequestMethod != null)
                {
                    OptionSetValue value = this.RequestMethod.Get<OptionSetValue>(context);
                    if (value != null && value.Value != 0)
                    {
                        switch (value.Value)
                        {
                            case GET:
                                response = client.GetAsync(RequestUrl.Get<string>(context));
                                break;
                            case PATCH:
                                response = client.PatchAsync(RequestUrl.Get<string>(context), content);
                                break;
                            case PUT:
                                response = client.PutAsync(RequestUrl.Get<string>(context), content);
                                break;
                            case DELETE:
                                response = client.DeleteAsync(RequestUrl.Get<string>(context));
                                break;
                            default:
                                response = client.PostAsync(RequestUrl.Get<string>(context), content);
                                break;

                        }
                    }
                }
                if (response != null)
                {
                    StringBuilder delimitedHeaders = new StringBuilder();
                    foreach (var header in response.Result.Headers)
                    {
                        if (delimitedHeaders.Length > 0)
                        {
                            delimitedHeaders.Append(";");
                        }
                        delimitedHeaders.Append($"{header.Key}:{header.Value}");
                    }
                    ResponseHeaders.Set(context, delimitedHeaders.ToString());
                    var responseString = response.Result.Content.ReadAsStringAsync();
                    ResponseBody.Set(context, responseString.Result);
                }
            }
        }
        public static string Stringify(EntitySpecification obj)
        {
            // Create a stream to serialize the object to.  
            var ms = new MemoryStream();

            // Serializer the User object to the stream.  
            var ser = new DataContractJsonSerializer(obj.GetType());
            ser.WriteObject(ms, obj);
            byte[] json = ms.ToArray();
            ms.Close();
            return Encoding.UTF8.GetString(json, 0, json.Length);
        }

    }

    public static class HttpClientExtensions
    {
        public static async Task<HttpResponseMessage> PatchAsync(this HttpClient client, string requestUrl, HttpContent iContent)
        {
            Uri requestUri = new Uri(requestUrl);
            var method = new HttpMethod("PATCH");
            var request = new HttpRequestMessage(method, requestUri)
            {
                Content = iContent
            };

            HttpResponseMessage response = new HttpResponseMessage();
            return await client.SendAsync(request);
        }
    }
}
