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
    class KnigaParse
    {
        public MyList<List<string>> goParse_kniga(List<string> richText, MyList<List<string>> res, Label label, Label label1,
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

            // http://www.kniga.ru/search/?search_query=9785699782703

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

                string site_str = "http://www.kniga.ru/search/?search_query=" + isbn;

                var page = Responses.GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                //HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                //HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                if (page != null)
                    doc.LoadHtml(page); // вся страница
                else
                    break;

                //<meta property="og:url" content="," />
                //link = "http://read.ru/id/";
                //link += GetWithInChString(page, "data-book-id=\"", '"');
                //link += "/";
                link = site_str;

                HtmlNodeCollection NodeLink = doc.DocumentNode.SelectNodes(".//div[@class='subNav']/a"); // name
                if (NodeLink != null)
                {
                    // string link = n.GetAttributeValue("href", ""); // получем ссылку
                    link = "http://www.kniga.ru";
                    link += NodeLink[0].GetAttributeValue("href", "");
                    page = Responses.GetPageCode(link, proxy);

                    if (page != null)
                        doc.LoadHtml(page); // вся страница

                    HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1"); // 
                    if (NodeName != null)
                    {
                        //if (NodeAuthor[0].InnerText.Contains(":"))
                        name = NodeName[0].InnerText;
                    }

                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//span[@id='authorsList']/a"); // 
                    if (NodeAuthor != null)
                    {
                        author = NodeAuthor[0].InnerText;
                    }

                    HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[2]/a"); // 
                    if (NodePublisher != null)
                    {
                        publisher = NodePublisher[0].InnerText;
                    }
                    else
                    {
                        HtmlNodeCollection NodePublisher2 = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[1]/a"); // 
                        if (NodePublisher2 != null)
                        {
                            publisher = NodePublisher2[0].InnerText;
                        }
                    }

                    HtmlNodeCollection NodeYear = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[2]/span[2]"); // 
                    if (NodeYear != null)
                    {
                        year = NodeYear[0].InnerText;
                        StringBuilder sb = new StringBuilder(year.Length);
                        foreach (char ch in year)
                        {
                            if (((ch >= '0' && ch <= '9')))
                            {
                                sb.Append(ch);
                            }
                        }
                        year = sb.ToString();// собранная строка
                    }

                    // .//*[@id='normalPrice']
                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//*[@id='normalPrice']/span"); // price
                    string str_price = "";
                    if (NodePrice != null)
                    {
                        price = NodePrice[0].InnerText;
                        StringBuilder sb = new StringBuilder(price.Length);

                        foreach (char ch in price)
                        {
                            if (((ch >= '0' && ch <= '9')) || ch == '.')
                            {
                                sb.Append(ch);
                            }
                        }
                        str_price = sb.ToString().Split('.')[0];// собранная строка
                    }
                    else
                    {
                        NodePrice = doc.DocumentNode.SelectNodes(".//*[@id='normalPrice']"); // price
                        if (NodePrice != null)
                        {
                            str_price = NodePrice[0].InnerText;
                        }
                    }

                    List<string> res_str = new List<string>();
                    res_str.Add(link);
                    res_str.Add(isbn);
                    res_str.Add(name);
                    res_str.Add(author);
                    res_str.Add(str_price);
                    res_str.Add(publisher);
                    res_str.Add(year);
                    res.Add(res_str);
                    i_label++;
                    label1.RefreshData((i_label).ToString());
                }
            }
            t1.Stop();
            return res;
        }
    }
}
