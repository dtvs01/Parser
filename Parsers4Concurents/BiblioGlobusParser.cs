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
    class BiblioGlobusParser
    {
        public MyList<List<string>> goParse_biblio_globus(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            string name = "";
            string isbn = "";
            int i_label = 0;
            //http://www.biblio-globus.ru/service/catalog/products?query=%d2%e0%ed%f6%f3%fe%f9%e0%ff+%f1+%eb%ee%f8%e0%e4%fc%ec%e8+(%cc%ee%e9%e5%f1)
            //http://www.biblio-globus.ru/service/catalog/products?query=%D0%A2%D0%B0%D0%BD%D1%86%D1%83%D1%8E%D1%89%D0%B0%D1%8F%20%D1%81%20%D0%BB%D0%BE%D1%88%D0%B0%D0%B4%D1%8C%D0%BC%D0%B8%20%28%D0%9C%D0%BE%D0%B9%D0%B5%D1%81%29
            //string site_str = "http://www.100000-knig.ru/catalog/search?author=%E0%E2%F2%EE%F0&name=" + stringOfSearchGo + "&publisher=%E8%E7%E4%E0%F2%E5%EB%FC%F1%F2%E2%EE&Search=%C8%F1%EA%E0%F2%FC";

            for (int i = 0; i < richText.Count(); i++)
            {
                label.RefreshData((i + 1).ToString());
                string[] name_isbn = richText[i].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];

                string[] strFullName = name.Split(' ');
                int len = strFullName.Count();
                if (name.Length > 90)
                {
                    name = name.Substring(0, 90);
                }

                string book_id = Responses.POST("http://www.bgshop.ru/Default.aspx", "ucSearchBrief$txtSearch=" + isbn);
                // разборки с кодировкой !!!
                //Encoding srcEncodingFormat = Encoding.UTF8;
                //Encoding dstEncodingFormat = Encoding.GetEncoding("windows-1251");
                //byte[] originalByteString = srcEncodingFormat.GetBytes(name);
                //byte[] convertedByteString = Encoding.Convert(srcEncodingFormat,
                //dstEncodingFormat, originalByteString);

                string site_str = "http://www.biblio-globus.ru/service/catalog/products?query=";
                string strSearch = site_str + HttpUtility.UrlEncode(name);

                while ((len > 1))
                {
                    var page = Responses.GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(page); // вся страница

                    // здесь и производится выборка интересующих нодов
                    // В данном случае выбираем блочные элементы с классом eTitle div class="BookAttrs"
                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes(".//a[@class='name']");
                    if (NodeTR != null)
                    {
                        foreach (HtmlNode n in NodeTR)
                        {
                            string link = n.GetAttributeValue("href", ""); // получем ссылку
                            var page_detail = Responses.GetPageCode("http://www.biblio-globus.ru" + link);
                            doc_isbn.LoadHtml(page_detail);
                            HtmlNode NodeISBN = doc_isbn.DocumentNode.SelectSingleNode(".//div[@class='details']");

                            string txt_isbn = NodeISBN.InnerHtml;
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

                                // фильтр по ISBN-10 && ISBN-13 
                                string isbn_l = isbn.Substring(3, 10);
                                if (isbn == str_isbn || isbn_l == str_isbn)
                                {
                                    HtmlNode NodePrice = doc_isbn.DocumentNode.SelectSingleNode(".//div[@class='details_price']");
                                    string price = NodePrice.InnerText; // получем ссылк

                                    // чистим price
                                    StringBuilder sb2 = new StringBuilder(price.Length);
                                    foreach (char ch in price)
                                    {
                                        if ((ch >= '0' && ch <= '9') || ch == ',')
                                        {
                                            sb2.Append(ch);
                                        }
                                    }
                                    string str_price = sb2.ToString(); // собранная строка

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
                    strSearch = strSearch.Substring(0, last);
                    len -= 1;
                }
            }
            t1.Stop();
            return res;
        }
    }
}
