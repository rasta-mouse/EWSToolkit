using Microsoft.Exchange.WebServices.Data;

using System.Collections.Generic;

namespace Toolkit
{
    public class Search
    {
        ExchangeService Service;
        List<SearchFilter> SearchFilters;
        int SearchLimit;

        public Search(ExchangeConnection connection, int searchLimit = 50)
        {
            Service = connection.ExchangeService;
            SearchLimit = searchLimit;
            SearchFilters = new List<SearchFilter>();
        }

        public void AddSubjectSearchTerms(string[] terms)
        {
            foreach (var term in terms)
            {
                AddSubjectSearchTerm(term);
            }
        }

        public void AddSubjectSearchTerm(string term)
        {
            SearchFilters.Add(new SearchFilter.ContainsSubstring(ItemSchema.Subject, term));
        }

        public void AddBodySearchTerms(string[] terms)
        {
            foreach (var term in terms)
            {
                AddBodySearchTerm(term);
            }
        }

        public void AddBodySearchTerm(string term)
        {
            SearchFilters.Add(new SearchFilter.ContainsSubstring(ItemSchema.Body, term));
        }

        public List<SearchResult> SearchFolder(FolderName folder, bool returnBody = true, bool returnAttachments = true)
        {
            var searchResult = new List<SearchResult>();

            var searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.Or, SearchFilters.ToArray());

            var view = new ItemView(SearchLimit)
            {
                PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.HasAttachments, ItemSchema.DateTimeReceived)
            };

            view.OrderBy.Add(ItemSchema.DateTimeReceived, SortDirection.Descending);
            view.Traversal = ItemTraversal.Shallow;

            var res = Service.FindItems((WellKnownFolderName)folder, searchFilter, view);

            if (res.TotalCount > 0)
            {
                foreach (var item in res.Items)
                {
                    if (item is EmailMessage)
                    {
                        searchResult.Add(new SearchResult {
                            UniqueId = item.Id.UniqueId,
                            ReceivedOn = item.DateTimeReceived,
                            Subject = item.Subject,
                            HasAttachments = item.HasAttachments
                        });
                    }
                }
            }

            return searchResult;
        }
    }
}