using Microsoft.Exchange.WebServices.Data;

namespace Toolkit
{
    public class EmailRule
    {
        ExchangeService Service;
        Rule Rule;

        public EmailRule(ExchangeConnection connection)
        {
            Service = connection.ExchangeService;
            Rule = new Rule();
        }

        public void SetDisplayName(string name)
        {
            Rule.DisplayName = name;
        }

        public void AddSubjectConditions(string[] keywords)
        {
            foreach (var keyword in keywords)
            {
                AddSubjectCondition(keyword);
            }
        }

        public void AddSubjectCondition(string keyword)
        {
            Rule.Conditions.ContainsSubjectStrings.Add(keyword);
        }

        public void AddBodyConditions(string[] keywords)
        {
            foreach (var keyword in keywords)
            {
                AddBodyCondition(keyword);
            }
        }

        public void AddBodyCondition(string keyword)
        {
            Rule.Conditions.ContainsBodyStrings.Add(keyword);
        }

        public void AddSenderConditions(string[] senders)
        {
            foreach (var sender in senders)
            {
                AddSenderCondition(sender);
            }
        }

        public void AddSenderCondition(string sender)
        {
            Rule.Conditions.ContainsSenderStrings.Add(sender);
        }

        public void AddForwardAction(string forwardAddress)
        {
            Rule.Actions.ForwardAsAttachmentToRecipients.Add(forwardAddress);
        }

        public void AddDeleteAction(bool permanent = true)
        {
            Rule.Actions.Delete = true;

            if (permanent)
            {
                Rule.Actions.PermanentDelete = true;
            }
        }

        public void AddToInbox()
        {
            var rule = new CreateRuleOperation(Rule);
            Service.UpdateInboxRules(new RuleOperation[] { rule }, true);
        }
    }
}