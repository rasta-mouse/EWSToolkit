using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Exchange.WebServices.Data;

namespace EWS
{
    public class Rules
    {
        public static void AddNewRule(ExchangeService service, string name, List<string> subjects, List<string> bodies, string forward)
        {
            Console.WriteLine(" [>] Constructing rule");

            Rule newRule = new Rule();

            // set name

            Console.WriteLine(" [>] Name: {0}", name);
            newRule.DisplayName = name;

            // set subjects
            if (subjects.Any())
            {
                foreach (string subject in subjects)
                {
                    Console.WriteLine(" [>] Subject: {0}", subject);
                    newRule.Conditions.ContainsSubjectStrings.Add(subject);
                }
            }

            // set body
            if (bodies.Any())
            {
                foreach (string body in bodies)
                {
                    Console.WriteLine(" [>] Body: {0}", body);
                    newRule.Conditions.ContainsBodyStrings.Add(body);
                }
            }

            // forwartd to
            Console.WriteLine(" [>] Forward to: {0}", forward);
            newRule.Actions.ForwardAsAttachmentToRecipients.Add(forward);

            Console.WriteLine(" [>] Adding rule");

            CreateRuleOperation createRule = new CreateRuleOperation(newRule);
            service.UpdateInboxRules(new RuleOperation[] { createRule }, true);

            Console.WriteLine(" [>] Done");

        }
    }
}
