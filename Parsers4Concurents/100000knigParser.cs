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
    class _100000knigParser
    {
        public MyList<List<string>> goParse100000(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            string name = "";
            string isbn = "";
            int i_label = 0;

            //string site_str = "http://www.100000-knig.ru/catalog/search?author=%E0%E2%F2%EE%F0&name=" + stringOfSearchGo + "&publisher=%E8%E7%E4%E0%F2%E5%EB%FC%F1%F2%E2%EE&Search=%C8%F1%EA%E0%F2%FC";
            int brk = 0;
            for (int i = 0; i < richText.Count(); i++)
            {

                if (brk > 0)
                    return res;

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

                // разборки с кодировкой !!!
                Encoding srcEncodingFormat = Encoding.UTF8;
                Encoding dstEncodingFormat = Encoding.GetEncoding("windows-1251");
                byte[] originalByteString = srcEncodingFormat.GetBytes(name);
                byte[] convertedByteString = Encoding.Convert(srcEncodingFormat,
                dstEncodingFormat, originalByteString);

                string site_str = "http://www.100000-knig.ru/catalog/search?author=%E0%E2%F2%EE%F0&name=";
                string strSearch = site_str + HttpUtility.UrlEncode(convertedByteString) + "&publisher=%E8%E7%E4%E0%F2%E5%EB%FC%F1%F2%E2%EE&Search=%C8%F1%EA%E0%F2%FC";

                while ((len > 1))
                {
                    var page = Responses.GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                    if (page != null)
                    {
                        doc.LoadHtml(page); // вся страница
                    }
                    else
                    {
                        brk++;
                        break;
                    }

                    // здесь и производится выборка интересующих нодов
                    // В данном случае выбираем блочные элементы с классом eTitle div class="BookAttrs"
                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//div[@class='smallbook-zakaz']");
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

                                // фильтр по ISBN-10 && ISBN-13 
                                string isbn_l = isbn.Substring(3, 10);
                                if (isbn == str_isbn || isbn_l == str_isbn)
                                {
                                    doc_isbn.LoadHtml(n.InnerHtml);
                                    HtmlNode NodeHref = doc_isbn.DocumentNode.SelectSingleNode("//div[@class='tov-button']/a");
                                    string href = NodeHref.GetAttributeValue("href", ""); // получем ссылку

                                    // получаем страницу по ссылке
                                    var page_price = Responses.GetPageCode("http://www.100000-knig.ru" + href);
                                    doc_price.LoadHtml(page_price);
                                    HtmlNode NodePrice = doc_price.DocumentNode.SelectSingleNode("//div[@id='tovar-price']");
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
                                    string str_price = sb2.ToString();// собранная строка
                                    string[] pr = str_price.Split(',');
                                    str_price = pr[0];

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
