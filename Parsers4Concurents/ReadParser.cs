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
    class ReadParser
    {
        public MyList<List<string>> goParse_read(List<string> richText, MyList<List<string>> res, Label label, Label label1,
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

            //http://read.ru/search/?search=9785699792788

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

                string site_str = "http://read.ru/search/?search=" + isbn;

                var page = Responses.GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                //HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                //HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                if (page != null)
                    doc.LoadHtml(page); // вся страница

                //<meta property="og:url" content="," />
                link = "http://read.ru/id/";
                link += Strings.GetWithInChString(page, "data-book-id=\"", '"');
                link += "/";

                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1[@class='book_card__header']"); // name 
                if (NodeName != null)
                {
                    name = NodeName[0].InnerText;
                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='j-book_autors book_autors']"); // 
                    if (NodeAuthor != null)
                    {
                        if (NodeAuthor[0].InnerText.Contains(":"))
                            author = NodeAuthor[0].InnerText.Split(':')[1].Trim();
                    }

                    HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//li[@class='book_properties__item j-book_pub']"); // 
                    if (NodePublisher != null)
                    {
                        publisher = NodePublisher[0].InnerText.Split(':')[1].Trim();
                    }

                    // book_price3__title_off
                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//div[@class='book_price3__title_off']"); // price
                    if (NodePrice != null)
                    {
                        price = NodePrice[0].InnerText;
                    } // book_price3__fullprice
                    else
                    {
                        // floatLeft
                        NodePrice = doc.DocumentNode.SelectNodes(".//div[@class='book_price3__fullprice']/div[@class='floatLeft']"); // price
                        if (NodePrice != null)
                            price = NodePrice[0].InnerText;
                    }

                    // li book_properties__item-promo // span book_field
                    HtmlNodeCollection NodeYear = doc.DocumentNode.SelectNodes(".//td[@class='year']/span[1]"); // 
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
            }
            t1.Stop();
            return res;
        }
    }
}
