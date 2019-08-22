using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json;
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
        public string EntityLogicalName { get; set; }
        public Guid EntityId { get; set; }
    }
    public class WebHook : ContributingClasses.WorkflowBase
    {
        [Input("RequestHeaders")]
        public InArgument<string> RequestHeaders { get; set; }

        [RequiredArgument]
        [Input("Request Url")]
        public InArgument<string> RequestUrl { get; set; }

        [Input("Send Current Record as Request Body")]
        public InArgument<bool> SendCurrentRecordAsBody { get; set; }

        [Input("Request Body")]
        public InArgument<string> RequestBody { get; set; }


        public const int GET = 222540000;
        public const int HEAD = 222540001;
        public const int POST = 222540002;
        public const int PUT = 222540003;
        public const int DELETE = 222540004;
        public const int TRACE = 222540005;
        public const int OPTIONS = 222540006;
        public const int CONNECT = 222540007;
        public const int PATCH = 222540008;

        public const int ASYNC = 222540001;

        [RequiredArgument]
        [Input("Synchronous Mode")]
        [AttributeTarget("afk_workflowelementoption", "afk_synchronousmode")]
        public InArgument<OptionSetValue> SynchronousMode { get; set; }

        [RequiredArgument]
        [Input("Request Method")]
        [AttributeTarget("afk_workflowelementoption", "afk_httpverb")]
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
                if (SendCurrentRecordAsBody.Get(context))
                {
                    var entity = (Entity)workflowContext.InputParameters["Target"];
                    body = JsonConvert.SerializeObject(entity);
                }
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
                        response = SendRequest(context, client, value.Value, content);
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
        public async Task<HttpResponseMessage> SendRequest(CodeActivityContext context, HttpClient client, int method, StringContent content)
        {
            OptionSetValue syncMode = this.SynchronousMode.Get<OptionSetValue>(context);
            if (syncMode != null && syncMode.Value == ASYNC)
            {
                switch (method)
                {
                    case GET:
                        client.GetAsync(RequestUrl.Get<string>(context));
                        return null;
                    case PATCH:
                        client.PatchAsync(RequestUrl.Get<string>(context), content);
                        return null;
                    case PUT:
                        client.PutAsync(RequestUrl.Get<string>(context), content);
                        return null;
                    case DELETE:
                        client.DeleteAsync(RequestUrl.Get<string>(context));
                        return null;
                    default:
                        client.PostAsync(RequestUrl.Get<string>(context), content);
                        return null;
                }

            }
            else
            {
                switch (method)
                {
                    case GET:
                        return await client.GetAsync(RequestUrl.Get<string>(context));
                    case PATCH:
                        return await client.PatchAsync(RequestUrl.Get<string>(context), content);
                    case PUT:
                        return await client.PutAsync(RequestUrl.Get<string>(context), content);
                    case DELETE:
                        return await client.DeleteAsync(RequestUrl.Get<string>(context));
                    default:
                        return await client.PostAsync(RequestUrl.Get<string>(context), content);
                }
            }
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
