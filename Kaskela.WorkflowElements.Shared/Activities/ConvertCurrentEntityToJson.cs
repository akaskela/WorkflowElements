using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Workflow;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
    public class SerializeCurrentEntity : ContributingClasses.WorkflowBase
    {
        [Output("Stringified Entity")]
        public OutArgument<string> SerializedEntity { get; set; }

        protected override void Execute(CodeActivityContext context)
        {
            var workflowContext = context.GetExtension<IWorkflowContext>();
            var entity = (Entity)workflowContext.InputParameters["Target"];
            SerializedEntity.Set(context, JsonConvert.SerializeObject(entity));
        }
    }
}
