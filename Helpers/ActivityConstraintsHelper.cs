using System;
using System.Activities;
using System.Activities.Statements;
using System.Activities.Validation;
using System.Collections.Generic;
using System.Linq;

namespace Bysxiang.UipathExcelEx.Helpers
{
    public static class ActivityConstraintsHelper
    {
        public static Constraint GetCheckParentConstraint<ActivityType>(string parentTypeName, string validationMessage = null) 
            where ActivityType : Activity
        {
            return GetCheckParentConstraint<ActivityType>(new string[] { parentTypeName }, validationMessage);
        }

        public static Constraint GetCheckParentConstraint<ActivityType>(string[] parentTypeNames, string validationMessage)
            where ActivityType : Activity
        {
            DelegateInArgument<ValidationContext> delegateInArgument = new DelegateInArgument<ValidationContext>();
            DelegateInArgument<ActivityType> argument = new DelegateInArgument<ActivityType>();
            DelegateInArgument<Activity> parent = new DelegateInArgument<Activity>();
            Variable<bool> variable = new Variable<bool>();
            Variable<IEnumerable<Activity>> variable2 = new Variable<IEnumerable<Activity>>();
            Constraint<ActivityType> constraint = new Constraint<ActivityType>();
            constraint.Body = new ActivityAction<ActivityType, ValidationContext>
            {
                Argument1 = argument,
                Argument2 = delegateInArgument,
                Handler = new Sequence
                {
                    Variables =
                    {
                        variable,
                        variable2
                    },
                    Activities =
                    {
                        new Assign<IEnumerable<Activity>>
                        {
                            To = variable2,
                            Value = new GetParentChain
                            {
                                ValidationContext = delegateInArgument
                            }
                        },
                        new ForEach<Activity>
                        {
                            Values = variable2,
                            Body = new ActivityAction<Activity>
                            {
                                Argument = parent,
                                Handler = new If
                                {
                                    Condition = new InArgument<bool>((ActivityContext ctx) => parentTypeNames.Contains(parent.Get(ctx).GetType().Name)),
                                    Then = new Assign<bool>
                                    {
                                        Value = true,
                                        To = variable
                                    }
                                }
                            }
                        },
                        new AssertValidation
                        {
                            Assertion = new InArgument<bool>(variable),
                            Message = new InArgument<string>(validationMessage)
                        }
                    }
                }
            };

            return constraint;
        }
    }
}
