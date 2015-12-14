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
    class ProdalitParser
    {
        public MyList<List<string>> goParse_prodalit(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            string name = "";
            string isbn = "";
            int i_label = 0;
            string site_str = "http://www.prodalit.ru/asp/cat.aspx?Mode=find&FormShort=1&Txt=";

            for (int i = 0; i < richText.Count(); i++)
            {
                label.RefreshData((i + 1).ToString());
                string[] name_isbn = richText[i].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];

                string[] strFullName = name.Split(' ');
                int len = strFullName.Count();
                if (name.Length > 70)
                {
                    name = name.Substring(0, 70);
                }

                string strSearch = site_str + HttpUtility.UrlEncode(name); //HttpUtility.UrlEncode(convertedByteString)
                while ((len > 1))
                {
                    var page = Responses.GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page); // вся страница

                    // здесь и производится выборка интересующих нодов
                    // В данном случае выбираем блочные элементы с классом eTitle div class="BookAttrs"
                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//div[@class='BookAttrs']");
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
                                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes("//div[@class='ShopCostAction']");
                                    string price_booka = "";
                                    if (NodePrice != null)
                                    {
                                        price_booka = NodePrice[0].InnerText;
                                    }
                                    else
                                    {
                                        HtmlNodeCollection NodePriceNoAction = doc.DocumentNode.SelectNodes("//span[@class='ShopCost']");
                                        if (NodePriceNoAction != null)
                                        {
                                            price_booka = NodePriceNoAction[0].InnerText;
                                        }
                                    }

                                    // чистим price
                                    StringBuilder sb2 = new StringBuilder(price_booka.Length);
                                    foreach (char ch in price_booka)
                                    {
                                        if ((ch >= '0' && ch <= '9') || ch == ',')
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
                                    i_label++;
                                    label1.RefreshData((i_label).ToString());
                                    len = 0;
                                }
                            }
                        }
                    }

                    int last = strSearch.LastIndexOf('+');
                    if (last > 0)
                        strSearch = strSearch.Substring(0, last);
                    len -= 1;
                }
            }
            t1.Stop();
            return res;
        }
    }
}
