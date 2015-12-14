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
    class OmskParser
    {
        int i_omsk = 0; // переменная цикла парсера omskbook
        public MyList<List<string>> goParse_omskbook(List<string> richText, MyList<List<string>> res, Label label, Label label1, System.Windows.Forms.Timer t1)
        {
            //string strCSV = "";
            string name = "";
            string isbn = "";
            string price = "";
            int i_str = 1;
            int i_label = 0;

        BR:
            for (; i_omsk < richText.Count(); i_omsk++)
            {
                label.RefreshData((i_str + i_omsk).ToString());
                string[] name_isbn = richText[i_omsk].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];

                //string[] strFullName = name.Split(' ');

                for (int j = 0; j < 4; j++)
                {
                    string site_str = "http://www.omskbook.ru/catalog/?page=";

                    // разборки с кодировкой !!!
                    //string str = "UTF8 Encoded string.";
                    Encoding srcEncodingFormat = Encoding.UTF8;
                    Encoding dstEncodingFormat = Encoding.GetEncoding("windows-1251");
                    byte[] originalByteString = srcEncodingFormat.GetBytes(name);
                    byte[] convertedByteString = Encoding.Convert(srcEncodingFormat,
                    dstEncodingFormat, originalByteString);
                    //string final_name = dstEncodingFormat.GetString(convertedByteString);

                    string strSearch = site_str + j.ToString() + "&search=" + HttpUtility.UrlEncode(convertedByteString);

                    string page = null;
                    page = Responses.GetPageCode(strSearch);
                    if (page != null)
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(page); // вся страница

                        HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//table[@width='670']//tr");
                        if (NodeTR != null)
                        {
                            foreach (HtmlNode n in NodeTR)
                            {
                                // Получаем строчки
                                // фрагмент который может содержать isbn 
                                doc_isbn.LoadHtml(n.InnerHtml);
                                HtmlNodeCollection NodeISBN = doc_isbn.DocumentNode.SelectNodes("//span[@class='style_2nd']");
                                //HtmlNodeCollection NodeISBN = doc_isbn.DocumentNode.SelectNodes("//tr");
                                if (NodeISBN != null)
                                {
                                    foreach (HtmlNode isbn_n in NodeISBN)
                                    {
                                        //Получаем строчки
                                        string[] txt = isbn_n.InnerText.Split(':');
                                        string str = txt[4].Trim().Replace("-", string.Empty);
                                        // если isbn совпали, пишем...
                                        if (str == isbn)
                                        {
                                            //HtmlNodeCollection NodePrice = doc_isbn.DocumentNode.SelectNodes("//tr[1]/td[3]/span"); style_name
                                            HtmlNodeCollection NodePrice = doc_isbn.DocumentNode.SelectNodes("//td[@align='center']//span[@class='style_name']");
                                            if (NodePrice != null)
                                            {
                                                foreach (HtmlNode n_price in NodePrice)
                                                {
                                                    //Получаем строчки
                                                    price = n_price.InnerText;

                                                    // преобразуем цену
                                                    StringBuilder sb = new StringBuilder(price.Length);
                                                    foreach (char ch in price)
                                                    {
                                                        if ((ch >= '0' && ch <= '9') || ch == '.')
                                                        {
                                                            sb.Append(ch);
                                                        }
                                                    }
                                                    price = sb.ToString();// собранная строка
                                                    price = price.Trim(new char[] { '.' }); // обрезка последней точки

                                                    //string result = isbn + ";" + name + ";" + price + ";";
                                                    List<string> res_str = new List<string>();
                                                    res_str.Add(isbn);
                                                    res_str.Add(name);
                                                    res_str.Add(price);
                                                    res.Add(res_str);
                                                    i_label++;
                                                    label1.RefreshData((i_label).ToString());
                                                    i_omsk++; // принудительное прерывание цикла
                                                    goto BR;
                                                }
                                            }

                                        }

                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        //MessageBox.Show("По не запросу получена страница!");
                    }
                }
            }
            t1.Stop();
            return res;
        }
    }
}
