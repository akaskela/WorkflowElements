using Kaskela.WorkflowElements.Shared.ContributingClasses;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace Kaskela.WorkflowElements.Shared.Activities
{
    public class DateAddOrSubtract : WorkflowBase
    {
        [RequiredArgument]
        [Input("Date to Modify")]
        public InArgument<DateTime> DateToModify { get; set; }

        [RequiredArgument]
        [Input("Unit")]
        [AttributeTarget("afk_workflowelementoption", "afk_datetimeunits")]
        public InArgument<OptionSetValue> Units { get; set; }

        [RequiredArgument]
        [Input("Number")]
        public InArgument<int> Number { get; set; }
        
        [Output("Modified Date")]
        public OutArgument<DateTime> ModifiedDate { get; set; }
        protected override void Execute(CodeActivityContext context)
        {
            bool dateSet = false;
            if (this.Units != null && this.Number != null && this.Number.Get<int>(context) != 0)
            {
                OptionSetValue value = this.Units.Get<OptionSetValue>(context);
                if (value != null)
                {
                    switch (value.Value)
                    {
                        case 222540000:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddYears(this.Number.Get<int>(context)));
                            break;
                        case 222540001:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddMonths(this.Number.Get<int>(context)));
                            break;
                        case 222540002:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddDays(this.Number.Get<int>(context) * 7));
                            break;
                        case 222540003:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddDays(this.Number.Get<int>(context)));
                            break;
                        case 222540004:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddHours(this.Number.Get<int>(context)));
                            break;
                        default:
                            ModifiedDate.Set(context, this.DateToModify.Get(context).AddMinutes(this.Number.Get<int>(context)));
                            break;
                    }
                    dateSet = true;
                }
            }

            if (!dateSet)
            {
                ModifiedDate.Set(context, this.DateToModify.Get(context));
            }
        }
    }
}
