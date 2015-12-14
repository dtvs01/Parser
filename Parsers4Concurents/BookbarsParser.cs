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
    class BookbarsParser
    {
        public MyList<List<string>> goParse_bookbars(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            string name = "";
            string isbn = "";
            string link = "";
            string publisher = "";
            string author = "";
            string year = "";
            string price = "";
            int i_label = 0;

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

                //http://bookbars.ru/search/content/9785699782703
                string site_str = "http://bookbars.ru/search/content/" + isbn;

                var page = Responses.GetPageCode(site_str);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 

                if (page != null)
                {
                    doc.LoadHtml(page); // вся страница
                    link = "http://bookbars.ru" + Strings.GetWithInChString(page, "<h3><a href=\"", '"') + "/";
                }
                else
                {
                    //
                }


                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//div[@class='field-item even']/h3/a"); // name 
                if (NodeName != null)
                {
                    name = NodeName[0].InnerText;

                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='field field-name-field-kniga-author field-type-text field-label-hidden']/div/div"); // 
                    if (NodeAuthor != null)
                    {
                        author = NodeAuthor[0].InnerText.Trim();
                    }

                    HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//div[@class='field field-name-field-kniga-publisher field-type-text field-label-inline clearfix']/div[2]/div"); // 
                    if (NodePublisher != null)
                    {
                        publisher = NodePublisher[0].InnerText.Split(',')[0].Trim();
                    }

                    // book_price3__title_off
                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//div[@class='field field-name-commerce-price field-type-commerce-price field-label-hidden']/div/div"); // price
                    if (NodePrice != null)
                    {
                        price = NodePrice[0].InnerText.Split(' ')[0].Trim();
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

                    HtmlNodeCollection NodeYear = doc.DocumentNode.SelectNodes(".//div[@class='field field-name-field-kniga-year field-type-text field-label-inline clearfix']/div[2]/div"); // 
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
