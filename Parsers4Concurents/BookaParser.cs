using HtmlAgilityPack;
using Parsers4Concurents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace Parsers4Сompetitor
{
    class BookaParser
    {
        int i_booka = 0;
        public MyList<List<string>> goParse_booka(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            //string strCSV = "";
            string name = "";
            string isbn = "";
            int i_label = 0;

        BR: // рвем цикл постранички, если нашли значение
            for (; i_booka < richText.Count(); i_booka++)
            {
                label.RefreshData((i_booka + 1).ToString());
                string[] name_isbn = richText[i_booka].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];

                for (int j = 1; j < 4; j++) // постраничка
                {

                    string site_str = "http://www.booka.ru/search?page=" + j + "&st=title&q=";
                    string strSearch = site_str + HttpUtility.UrlEncode(name); //HttpUtility.UrlEncode(convertedByteString)

                    var page = Responses.GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page); // вся страница

                    // здесь и производится выборка интересующих нодов
                    // В данном случае выбираем блочные элементы с классом eTitle
                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//td[@class='content']");
                    if (NodeTR != null)
                    {
                        foreach (HtmlNode n in NodeTR)
                        {
                            string txt_isbn = n.InnerHtml;
                            if (txt_isbn.Contains("ISBN:"))
                            {
                                List<string> lst = Strings.GetWithIn(txt_isbn, "ISBN:", 28);
                                // чистим isbn
                                StringBuilder sb = new StringBuilder(lst[0].Length);
                                foreach (char ch in lst[0])
                                {
                                    if ((ch >= '0' && ch <= '9'))
                                    {
                                        sb.Append(ch);
                                    }
                                }
                                string str_isbn = sb.ToString();// собранная строка
                                if (isbn == str_isbn)
                                {
                                    doc_isbn.LoadHtml(n.InnerHtml);
                                    HtmlNodeCollection NodePrice = doc_isbn.DocumentNode.SelectNodes("//div[@class='sprice']");
                                    string price_booka = NodePrice[0].InnerText;

                                    // чистим price
                                    StringBuilder sb2 = new StringBuilder(price_booka.Length);
                                    foreach (char ch in price_booka)
                                    {
                                        if ((ch >= '0' && ch <= '9'))
                                        {
                                            sb2.Append(ch);
                                        }
                                    }
                                    string str_price = sb2.ToString();// собранная строка

                                    List<string> res_str = new List<string>();
                                    res_str.Add(isbn);
                                    res_str.Add(name);
                                    res_str.Add(str_price);
                                    res.Add(res_str);
                                    i_booka++;
                                    i_label++;
                                    label1.RefreshData((i_label).ToString());
                                    goto BR;
                                }

                            }
                        }
                    }
                }
            }
            t1.Stop();
            return res;
        }
    }
}
