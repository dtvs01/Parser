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
    class BiblionParser
    {
        public static MyList<List<string>> goParse_biblion(List<string> richText, MyList<List<string>> res, Label label, Label label1,
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

            // http://www.biblion.ru/search/?query=9785170865987&search=%CD%E0%E9%F2%E8

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

                string site_str = "http://www.biblion.ru/search/?query=" + isbn + "&search=%CD%E0%E9%F2%E8";

                var page = Responses.GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 

                if (page != null)
                    doc.LoadHtml(page); // вся страница

                link = "http://www.biblion.ru/product/";
                link += Strings.GetWithInChString(page, "inCart-", '"');
                link += "/";

                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//div[@class='product_details_text']/h2/a"); // name 
                if (NodeName != null)
                {
                    name = NodeName[0].InnerText;

                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='authors']/a"); // 
                    if (NodeAuthor != null)
                    {
                        author = NodeAuthor[0].InnerText.Trim();
                    }

                    HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//div[@class='shortDetails']/a"); // 
                    if (NodePublisher != null)
                    {
                        publisher = NodePublisher[0].InnerText.Trim();
                    }

                    // book_price3__title_off
                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//div[@class='right_block_product_details toCartPane']/div/span[1]"); // price
                    if (NodePrice != null)
                    {
                        price = NodePrice[0].InnerText;
                    }
                    /*
                    else
                    {
                        // floatLeft
                        NodePrice = doc.DocumentNode.SelectNodes(".//div[@class='book_price3__fullprice']/div[@class='floatLeft']"); // price
                        if (NodePrice != null)
                            price = NodePrice[0].InnerText;
                    }
                    */

                    HtmlNodeCollection NodeYear = doc.DocumentNode.SelectNodes(".//div[@class='shortDetails']/a[2]"); // 
                    if (NodeYear != null)
                    {
                        year = NodeYear[0].InnerText.Trim();
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
                    List<string> res_str = new List<string>();
                    res_str.Add(site_str);
                    res_str.Add(isbn);
                    res_str.Add(name);
                    res_str.Add("-");
                    res_str.Add("поиск не дал результатов");
                    res_str.Add("-");
                    res_str.Add("-");
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
