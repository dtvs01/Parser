using HtmlAgilityPack;
using Parsers4Concurents;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Parsers4Сompetitor
{
    class OzonParser
    {
        public MyList<List<string>> goParse_ozon(List<string> richText, MyList<List<string>> res, Label label, Label label1,
                                                    System.Windows.Forms.Timer t1, CheckBox checkTor,
                                                    string proxi = "", string login = "", string password = "")
        {
            string name = "";
            string isbn = "";
            string link = "";
            string publisher = "";
            string author = "";
            string year = "";
            int i_label = 0;

            //http://www.ozon.ru/
            //http://www.ozon.ru/?context=search&text=9785389077775&store=1,0 // 9785389072565
            //http://www.ozon.ru/?context=search&text=9785389072565&store=1,0 // подходит на оба случая

            TorClient Tor = null;
            if (checkTor.Checked)
            {
                Tor = new TorClient();
                Tor.StopTor();
                Tor.StartTor();
                Thread.Sleep(10000);
            }

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

                string site_str = "http://www.ozon.ru/?context=search&text=" + isbn + "&store=1,0";

                string page = null;
                if (Tor != null)
                {
                    page = Tor.RunParallel(site_str).Result;
                }
                else
                {
                    page = Responses.GetPageCode(site_str, proxi);
                }



                if (page != null)
                {
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                    HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();

                    doc.LoadHtml(page); // вся страница

                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes(".//div[@class='bAlsoSearch']"); // проверка страницы выдачи поиска
                    if (NodeTR != null)
                    {
                        // получаем ссылки
                        HtmlNodeCollection NodeBooks = doc.DocumentNode.SelectNodes(".//*[@id='bTilesModeShow']/div[1]");
                        if (NodeBooks != null)
                        {
                            foreach (HtmlNode n_boobs_href in NodeBooks)
                            {
                                string content_book = n_boobs_href.InnerHtml; // Нет в продаже
                                string str_price = "";
                                if (!content_book.Contains("Нет в продаже"))
                                {
                                    List<string> lst = Strings.GetWithIn(content_book, "\"price\" content=\"", 8);
                                    // берем цену
                                    StringBuilder sb = new StringBuilder(lst[0].Length);
                                    foreach (char ch in lst[0])
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
                                    str_price = "Нет в продаже";
                                }


                                //List<string> a_list = GetWithInCh(content_book, "href=\"", '"');
                                string href_book = Strings.GetWithInStr(content_book, "href=\"", "\">");
                                link = "http://www.ozon.ru" + href_book;

                                
                                string page_book = null;
                                if (Tor != null)
                                {
                                    page_book = Tor.RunParallel(link).Result;
                                }
                                else
                                {
                                    page_book = Responses.GetPageCode(link, proxi);
                                }


                                if (page_book != null)
                                    doc_book.LoadHtml(page_book);

                                HtmlNodeCollection NodeBook = doc_book.DocumentNode.SelectNodes(".//p[@itemprop='publisher']/a");
                                if (NodeBook != null)
                                {
                                    foreach (HtmlNode book in NodeBook)
                                    {
                                        publisher = book.InnerText;
                                    }
                                }

                                HtmlNodeCollection NodeAuthor = doc_book.DocumentNode.SelectNodes(".//p[@itemprop='author']/a");
                                if (NodeAuthor != null)
                                {
                                    foreach (HtmlNode auth in NodeAuthor)
                                    {
                                        author = auth.InnerText;
                                    }
                                }

                                HtmlNodeCollection NodeISBN = doc_book.DocumentNode.SelectNodes(".//p[@itemprop='isbn']");
                                if (NodeISBN != null)
                                {
                                    foreach (HtmlNode n_year in NodeISBN)
                                    {
                                        year = n_year.InnerText.Split(';')[1].Trim();
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
                    }
                    // если сразу книга
                    else
                    {
                        HtmlNodeCollection Node_detal_link = doc.DocumentNode.SelectNodes(".//div[@class='eSaleBlock_colorWrap']/h3"); //eSaleBlock_colorWrap

                        string link_id = Strings.GetWithInStr(doc.DocumentNode.InnerHtml, "s.products = \";", ";");
                        link = "http://www.ozon.ru/context/detail/id/" + link_id + "/";

                        if (Node_detal_link == null || !Node_detal_link[0].InnerHtml.Contains("Нет в продаже"))
                        {
                            HtmlNodeCollection Node_detal_book = doc.DocumentNode.SelectNodes("//div[@class='bDetailLogoBlock']");
                            if (Node_detal_book != null)
                            {
                                foreach (HtmlNode book in Node_detal_book)
                                {
                                    doc_book.LoadHtml(book.InnerHtml);
                                    break;
                                }

                                string price = "";

                                HtmlNodeCollection Node_author = doc_book.DocumentNode.SelectNodes("//p[@itemprop='author']");
                                if (Node_author != null)
                                {
                                    foreach (HtmlNode auth in Node_author)
                                    {
                                        author = auth.InnerText.Split(':')[1].Trim();
                                        break;
                                    }
                                }
                                HtmlNodeCollection Node_publisher = doc_book.DocumentNode.SelectNodes("//p[@itemprop='publisher']");
                                if (Node_publisher != null)
                                {
                                    foreach (HtmlNode pub in Node_publisher)
                                    {
                                        publisher = pub.InnerText.Split(':')[1].Trim();
                                        break;
                                    }
                                }
                                HtmlNodeCollection Node_price = doc.DocumentNode.SelectNodes("//div[@class='bOzonPrice']/span[1]");
                                if (Node_price != null)
                                {
                                    foreach (HtmlNode pr in Node_price)
                                    {
                                        price = pr.InnerText;
                                        break;
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
                        }
                        else
                        {
                            HtmlNodeCollection Node_detal_book = doc.DocumentNode.SelectNodes("//div[@class='bDetailLogoBlock']");
                            HtmlNodeCollection Node_author = doc_book.DocumentNode.SelectNodes("//p[@itemprop='author']");
                            if (Node_author != null)
                            {
                                foreach (HtmlNode auth in Node_author)
                                {
                                    author = auth.InnerText.Split(':')[1].Trim();
                                    break;
                                }
                            }
                            List<string> res_str = new List<string>();
                            res_str.Add(link);
                            res_str.Add(isbn);
                            res_str.Add(name);
                            res_str.Add(author);
                            res_str.Add("нет в продаже");
                            res_str.Add(publisher);
                            res_str.Add(year);
                            res.Add(res_str);
                            i_label++;
                            label1.RefreshData((i_label).ToString());
                        }
                    }
                }
                else
                    break;
            }
            t1.Stop();
            if(Tor != null)
                Tor.StopTor();
            return res;
        }
    }
}
