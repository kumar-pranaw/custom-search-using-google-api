using Google.Apis.Customsearch.v1;
using Google.Apis.Customsearch.v1.Data;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ScrapeLinkedinData
{
    public static class GoogleExtractor
    {
        public static List<ResultDataModel> ExtractCustomSearchData(string searchText)
        {
            const string apiKey = "AIzaSyCaYNeDcAN78yINlcFbBC37OjY4wXVyhSI";
            const string searchEngineId = "012726790682361633628:9coio4ce3q8";
            var customSearchService = new CustomsearchService(new BaseClientService.Initializer { ApiKey = apiKey });
            var listRequest = customSearchService.Cse.List(searchText);
            listRequest.Cx = searchEngineId;

            //List<Result> paging = new List<Result>();
            List<ResultDataModel> dataModel = new List<ResultDataModel>();
            var count = 0;
            while (dataModel != null)
            {
                Console.WriteLine($"Page {count}");
                listRequest.Start = count;
                listRequest.Execute().Items?.ToList().ForEach(x => dataModel.Add(new ResultDataModel
                {
                    Content = x.Snippet,
                    Link = x.Link,
                    Title = x.Title,
                    Name = x.HtmlTitle
                }));
                //if (paging != null)
                //    foreach (var item in paging)
                //        Console.WriteLine("Title : " + item.Title + Environment.NewLine + "Link : " + item.Link +
                //                          Environment.NewLine + Environment.NewLine);
                count++;
                if (count >= 10)
                {
                    break;
                }
            }
            return dataModel;
        }
    }
}
