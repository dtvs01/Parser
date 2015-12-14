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
    class MdkArbatParser
    {
        public MyList<List<string>> goParse_mdk_arbat(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            string name = "";
            string isbn = "";
            int i_label = 0;

            // http://mdk-arbat.ru/catalog?kw=%D2%E0%E9%ED%E0+%F1%E5%F0%E4%F6%E0

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

                // разборки с кодировкой !!!
                Encoding srcEncodingFormat = Encoding.UTF8;
                Encoding dstEncodingFormat = Encoding.GetEncoding("windows-1251");
                byte[] originalByteString = srcEncodingFormat.GetBytes(name);
                byte[] convertedByteString = Encoding.Convert(srcEncodingFormat,
                dstEncodingFormat, originalByteString);

                string site_str = "http://mdk-arbat.ru/catalog?kw=";
                string strSearch = site_str + HttpUtility.UrlEncode(convertedByteString);

                /*
               try
               {
                   using (IWebDriver driver = new ChromeDriver())
                   {
                       driver.Navigate().GoToUrl(strSearch);


                       // переход на карточку товара
                       IWebElement HomeToElement = driver.FindElement(By.ClassName("clik_div3"));
                       HomeToElement.Click();

                       // покупка
                       IWebElement ToBasket = driver.FindElement(By.XPath(@"//img[@src='/bitrix/templates/t1/images/buy_elem.png']"));
                       ToBasket.Click();


                       driver.Close();
                       driver.Dispose();
                   }
               }
               catch (Exception ex)
               {
                   MessageBox.Show(ex.Message);
               }
              */

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
                        //MessageBox.Show("Сервер долго не отдает страницу!");
                    }

                    // здесь и производится выборка интересующих нодов
                    // В данном случае выбираем блочные элементы с классом eTitle div class="BookAttrs"
                    HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//div[@class='good_description']");
                    if (NodeTR != null)
                    {
                        foreach (HtmlNode n in NodeTR)
                        {
                            string txt_isbn = n.InnerHtml;
                            doc_isbn.LoadHtml(txt_isbn);
                            HtmlNodeCollection NodeISBN = doc_isbn.DocumentNode.SelectNodes("//span[@class='isbn']");
                            foreach (HtmlNode n_isbn in NodeISBN)
                            {
                                txt_isbn = n_isbn.InnerHtml;
                            }
                            // чистим isbn
                            StringBuilder sb = new StringBuilder(txt_isbn.Length);
                            foreach (char ch in txt_isbn)
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

                                HtmlNode NodePrice = doc_isbn.DocumentNode.SelectSingleNode("//span[@class='price_info']");
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
