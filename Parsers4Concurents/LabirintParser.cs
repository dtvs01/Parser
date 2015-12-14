using HtmlAgilityPack;
using Parsers4Concurents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Parsers4Сompetitor
{
    class LabirintParser
    {
        public MyList<List<string>> goParse_labirint(List<string> richText, MyList<List<string>> res, Label label, Label label1,
                                                        System.Windows.Forms.Timer t1,
                                                        string proxy = "", string login = "", string password = "")
        {
            string name = "";
            string isbn = "";
            string link = "";
            string publisher = "";
            string author = "";
            string year = "";
            string price = "";
            int i_label = 0;

            //http://www.labirint.ru/search/9785699793129/?labsearch=1

            for (int i = 0; i < richText.Count(); i++)
            {
                label.RefreshData((i + 1).ToString());
                string[] name_isbn = richText[i].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];

                if (name.Contains("("))
                {
                    int last_author = name.LastIndexOf('(');
                    name = name.Substring(0, last_author);
                }

                string site_str = "http://www.labirint.ru/search/" + isbn + "/?labsearch=1";

                var page = Responses.GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 

                if (page != null)
                    doc.LoadHtml(page); // вся страница

                //<meta property="og:url" content="," />
                link = "http://www.labirint.ru/books/";
                link += Strings.GetWithInStr(page, "http://img1.labirint.ru/books/", "/big.jpg");
                link += "/";

                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1"); // name 
                if (NodeName != null)
                {
                    if (!NodeName[0].InnerText.Contains("у нас ничего нет"))
                    {
                        string[] arr_name = NodeName[0].InnerText.Split(':');
                        if (arr_name.Count() > 1)
                        {
                            name = arr_name[1].Trim();
                        }
                        else
                        {
                            name = arr_name[0].Trim();
                        }

                        HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//div[@class='publisher']"); // 
                        if (NodePublisher != null)
                        {
                            publisher = NodePublisher[0].InnerText.Split(':')[1].Split(',')[0].Trim();
                            year = NodePublisher[0].InnerText.Split(':')[1].Split(',')[1].Trim();
                        }

                        HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='authors']"); // 
                        if (NodeAuthor != null)
                        {
                            author = NodeAuthor[0].InnerText.Split(':')[1].Trim();
                        }

                        HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//span[@class='buying-pricenew-val-number']"); // price
                        if (NodePrice != null)
                        {
                            price = NodePrice[0].InnerText;
                        }
                        else
                        {
                            HtmlNodeCollection NodePriceOld = doc.DocumentNode.SelectNodes(".//span[@class='buying-price-val-number']"); // price
                            if (NodePriceOld != null)
                            {
                                price = NodePriceOld[0].InnerText;
                            }
                        }

                        List<string> res_str = new List<string>();
                        res_str.Add(link);
                        res_str.Add(isbn);
                        res_str.Add(name);
                        res_str.Add(author);
                        res_str.Add(price);
                        res_str.Add(publisher);
                        res_str.Add(year);
                        res.Add(res_str);
                        i_label++;
                        label1.RefreshData((i_label).ToString());
                    }
                    else
                    {
                        name = "товар не найден";
                        List<string> res_str = new List<string>();
                        res_str.Add(link);
                        res_str.Add(isbn);
                        res_str.Add(name);
                        res_str.Add(author);
                        res_str.Add(price);
                        res_str.Add(publisher);
                        res_str.Add(year);
                        res.Add(res_str);
                        i_label++;
                        label1.RefreshData((i_label).ToString());
                    }
                }
            }
            t1.Stop();
            return res;
        }
    }
}
