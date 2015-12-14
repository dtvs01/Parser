using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Web;
using System.Text.RegularExpressions;
using Parsers4Concurents;
using System.IO;
using System.Xml;
using System.Diagnostics;
using SharpCompress.Reader;
using SharpCompress.Common;
using System.Threading;
using System.Runtime.InteropServices;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace Parsers4Сompetitor
{
    class Request
    {
        public void LoadXML(string local_file_path, string link_file)
        {
            string fileName = local_file_path;//к примеру... файл.zip замените названием того что скачиваете
            if (File.Exists(fileName) != true)// если файла нет то просто скачиваем
            {
                WebClient client = new WebClient();
                //client.DownloadProgressChanged += new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
                //client.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
                client.DownloadFileAsync(new Uri(link_file), local_file_path);
            }
            else// если файл есть, удаляем и скачиваем новый
            {
                File.Delete(fileName);
                WebClient client = new WebClient();
                //client.DownloadProgressChanged += new DownloadProgressChangedEventHandler(client_DownloadProgressChanged);
                //client.DownloadFileCompleted += new AsyncCompletedEventHandler(client_DownloadFileCompleted);
                client.DownloadFileAsync(new Uri(link_file), local_file_path);
            }
        }


        public string GetPageCode(string strURL, string proxyAddress = "", string username = "", string password = "")
        {
            string PageHTML = null;
            try 
            {
                //HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create("http://wordstat.yandex.ru");

                // myHttpWebRequest.Headers.Add(HttpRequestHeader.Cookie, kyki); // вот здесь добавляем куки в myHttpWebRequest из WebBrowser
                //HttpWebResponse  myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                //создаём строку url,в которой будем указывать адресс сайта
                //string url = "https://www.google.com/analytics/web/?authuser=0#report/content-site-speed-overview/a25359375w74690838p77104458/";
                //Создаём объект , который будет выполнять запрос к URI(идентификатор ресурса)
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);

                //request.Proxy = new WebProxy("198.50.140.85:8089", false);

                //request.Accept = "*/*";
                //request.Headers.Add("Authorization", "OAuth 5a5b5820f3c24cb0a22ea3b432a51fa6");
                //request.Headers.Add("Accept-Encoding", "gzip, deflate");
                request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.2; WOW64; Trident/6.0; .NET4.0E; .NET4.0C; InfoPath.3; .NET CLR 3.5.30729; .NET CLR 2.0.50727; .NET CLR 3.0.30729)";
                request.AllowAutoRedirect = true;

                if(proxyAddress != "" && proxyAddress != null)
                {
                    WebProxy myProxy = new WebProxy(proxyAddress, true);
                    //Uri newUri = new Uri(proxyAddress);
                    //myProxy = (WebProxy)request.Proxy;
                    //myProxy.Address = newUri;

                    if(username != "" && password != "")
                    {
                        myProxy.Credentials = new NetworkCredential(username, password);
                    }                   
                    request.Proxy = myProxy;
                }
                

                //GetResponse - возвращает ответ на интернет-запрос
                HttpWebResponse response = null;
                response = (HttpWebResponse)request.GetResponse();
                if(response != null)
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(response.GetResponseStream(), Encoding.Default);
                    //Читаем поток от начала до конца
                    PageHTML = sr.ReadToEnd();
                    //закрываем поток
                    sr.Close();
                    return PageHTML;
                }
                else
                {
                    MessageBox.Show("response == null!");
                    return PageHTML;
                }

                //Реализуем считывание символов из потока байтов в определенной кодировке.
                //GetResponseStream -  возвращает поток данных из  интернет-ресурса
                //using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.Default))
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return PageHTML;
                //wc = WebClientProxi();
                //return outputText;
            }
            
        }

        public string GetPageCodeBG(string strURL)
        {
            string PageHTML = null;
            try
            {
                //HttpWebRequest myHttpWebRequest = (HttpWebRequest)HttpWebRequest.Create("http://wordstat.yandex.ru");

                // myHttpWebRequest.Headers.Add(HttpRequestHeader.Cookie, kyki); // вот здесь добавляем куки в myHttpWebRequest из WebBrowser
                //HttpWebResponse  myHttpWebResponse = (HttpWebResponse)myHttpWebRequest.GetResponse();

                //создаём строку url,в которой будем указывать адресс сайта
                //string url = "https://www.google.com/analytics/web/?authuser=0#report/content-site-speed-overview/a25359375w74690838p77104458/";
                //Создаём объект , который будет выполнять запрос к URI(идентификатор ресурса)
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(strURL);

                //request.Proxy = new WebProxy("198.50.140.85:8089", false);

                //request.Accept = "*/*";
                //request.Headers.Add("Authorization", "OAuth 5a5b5820f3c24cb0a22ea3b432a51fa6");
                //request.Headers.Add("Accept-Encoding", "gzip, deflate");
                request.UserAgent = "Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.2; WOW64; Trident/6.0; .NET4.0E; .NET4.0C; InfoPath.3; .NET CLR 3.5.30729; .NET CLR 2.0.50727; .NET CLR 3.0.30729)";
                request.AllowAutoRedirect = false;
                //GetResponse - возвращает ответ на интернет-запрос
                HttpWebResponse response = null;
                response = (HttpWebResponse)request.GetResponse();
                if (response != null)
                {
                    System.IO.StreamReader sr = new System.IO.StreamReader(response.GetResponseStream(), Encoding.Default);
                    //Читаем поток от начала до конца
                    PageHTML = sr.ReadToEnd();
                    //закрываем поток
                    sr.Close();
                    return PageHTML;
                }
                else
                {
                    MessageBox.Show("response == null!");
                    return PageHTML;
                }

                //Реализуем считывание символов из потока байтов в определенной кодировке.
                //GetResponseStream -  возвращает поток данных из  интернет-ресурса
                //using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.Default))


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return PageHTML;
                //wc = WebClientProxi();
                //return outputText;
            }

        }


        /// <summary>
        /// Отсылка POST-запроса на URL 
        /// </summary>
        /// <param name="Url"></param>
        /// <param name="Data"></param>
        /// <returns></returns>
        private string POST(string Url, string Data)
        {
            string Out = "";
            try
            {
                System.Net.WebRequest req = System.Net.WebRequest.Create(Url);
                req.Method = "POST";
                req.Timeout = 100000;
                req.ContentType = "application/x-www-form-urlencoded";
                byte[] sentData = Encoding.GetEncoding(1251).GetBytes(Data);
                req.ContentLength = sentData.Length;
                System.IO.Stream sendStream = req.GetRequestStream();
                sendStream.Write(sentData, 0, sentData.Length);
                sendStream.Close();
                System.Net.WebResponse res = req.GetResponse();
                System.IO.Stream ReceiveStream = res.GetResponseStream();
                System.IO.StreamReader sr = new System.IO.StreamReader(ReceiveStream, Encoding.UTF8);
                //Кодировка указывается в зависимости от кодировки ответа сервера
                Char[] read = new Char[256];
                int count = sr.Read(read, 0, 256);
                Out = String.Empty;
                while (count > 0)
                {
                    String str = new String(read, 0, count);
                    Out += str;
                    count = sr.Read(read, 0, 256);
                }
            }
            catch (Exception ex)
            {
               
            }

            return Out;
        }

        /*
        /// <summary>
        ///  по заданному прокси и url вытаскивает значения элементов (по xpath)
        /// </summary>
        /// <param name="url"></param>
        /// <param name="host"></param>
        /// <param name="node1"></param>
        /// <param name="node2"></param>
        /// <param name="node3"></param>
        /// <param name="node4"></param>
        /// <param name="node5"></param>
        /// <returns></returns>
        /// <summary>
        public List<string> GetPageAgility(string url, WebClient wc)
        {
            List<string> outputText = new List<string>();
            //WebClient wc = new WebClient();
            try
            {
                //WebClient wc = new WebClient();
                //wc.Proxy = new WebProxy(host, false);
                //wc = WebClientProxi(); GetPageCode(string strURL)
                //var page = wc.DownloadString(url);

                //var web = new HtmlWeb
                //{
                //    AutoDetectEncoding = false,
                //    OverrideEncoding = Encoding.Default,
                //};


                var page = GetPageCode(url);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml(page);

                //element = request.GetPageAgility(link, client, "//h1", "//div[@class='isbn']", "//span[@class='buying-price-val-number']", "//div[@class='publisher']", "//div[@id='product-title']/div/span"); 
                // Собственно, здесь и производится выборка интересующих нодов
                // В данном случае выбираем блочные элементы с классом eTitle
                HtmlNodeCollection Node1 = doc.DocumentNode.SelectNodes("//div[@class='isbn']"); //"//div[@class='isbn']"
                HtmlNodeCollection Node2 = doc.DocumentNode.SelectNodes("//h1");
                //  buying-pricenew-val-number  buying-priceold-val-number  
                HtmlNodeCollection Node_price = doc.DocumentNode.SelectNodes("//span[@class='buying-price-val-number']");
                HtmlNodeCollection Node_price_old = doc.DocumentNode.SelectNodes("//span[@class='buying-priceold-val-number']");
                HtmlNodeCollection Node_price_new = doc.DocumentNode.SelectNodes("//span[@class='buying-pricenew-val-number']");

                HtmlNodeCollection Node4 = doc.DocumentNode.SelectNodes("//div[@class='publisher']");
                HtmlNodeCollection Node5 = doc.DocumentNode.SelectNodes("//div[@id='product-title']/div/span");

                // проверка на наличие найденных узлов
                outputText.Add(url);

                if (Node1 != null)
                {
                    foreach (HtmlNode n in Node1)
                    {
                        //Получаем строчки
                        outputText.Add(n.InnerText);
                    }
                }
                else
                {
                    outputText.Add("нет значения");
                }

                if (Node2 != null)
                {
                    foreach (HtmlNode n in Node2)
                    {
                        //Получаем строчки
                        outputText.Add(n.InnerText);
                    }
                }
                else
                {
                    outputText.Add("нет значения");
                }


                string old_price = null;
                if (Node_price != null)
                {
                    foreach (HtmlNode n in Node_price)
                    {
                        //Получаем строчки
                        outputText.Add(n.InnerText);
                    }
                }

                if (Node_price_old != null)
                {
                    foreach (HtmlNode n in Node_price_old)
                    {
                        //Получаем строчки
                        old_price = n.InnerText;
                    }
                }

                if (Node_price_new != null)
                {
                    foreach (HtmlNode n in Node_price_new)
                    {
                        //Получаем строчки
                        outputText.Add(old_price + "/" + n.InnerText);
                    }
                }
                if (Node_price_new == null && Node_price_old == null && Node_price == null)
                {
                    outputText.Add("нет значения");
                }

                if (Node4 != null)
                {
                    foreach (HtmlNode n in Node4)
                    {
                        //Получаем строчки
                        outputText.Add(n.InnerText);
                    }
                }
                else
                {
                    outputText.Add("нет значения");
                }

                if (Node5 != null)
                {
                    foreach (HtmlNode n in Node5)
                    {
                        //Получаем строчки
                        outputText.Add(n.InnerText);
                    }
                }
                else
                {
                    outputText.Add("нет значения");
                }

                return outputText;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                //wc = WebClientProxi();
                return outputText;
            }
        }
        */

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
                    page = GetPageCode(strSearch);
                    if(page != null)
                    {
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(page); // вся страница

                        // здесь и производится выборка интересующих нодов
                        // В данном случае выбираем блочные элементы с классом eTitle table width="670"
                        //HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//tr");
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
                        MessageBox.Show("По не запросу получена страница!");
                    }
                }
            }
            t1.Stop();
            return res;
        }

        public MyList<List<string>> goParse_bookskazan(List<string> richText, MyList<List<string>> res, Label label1, Label label2, System.Windows.Forms.Timer t1)
        {

            var page = GetPageCode("http://www.bookskazan.ru/article/452/10/");

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(page); // вся страница

            HtmlNodeCollection NodeTR = doc.DocumentNode.SelectNodes("//table");
            string path_folder = Directory.GetCurrentDirectory();

            LoadXML(path_folder + "\\kazan1.rar", "http://www.bookskazan.ru/download/price/DomKnigi-hud.rar");
            
            try
            {
                Thread.Sleep(1500);
                using (Stream stream = File.OpenRead(path_folder + "\\kazan1.rar"))
                {
                    var reader = ReaderFactory.Open(stream);
                    while (reader.MoveToNextEntry())
                    {
                        if (!reader.Entry.IsDirectory)
                        {
                            reader.WriteEntryToDirectory(path_folder, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            DirectoryInfo dir = new DirectoryInfo(path_folder);
            var lst = new List<String>();

            Microsoft.Office.Interop.Excel.Application excelapp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Worksheet sheet = new Microsoft.Office.Interop.Excel.Worksheet();
            excelapp.FileValidation = Microsoft.Office.Core.MsoFileValidationMode.msoFileValidationSkip; // Так все работает

            foreach (FileInfo file in dir.GetFiles("*ДомКниги*"))
            {
                
                excelapp.Visible = false;
                //excelapp.UserControl = true;
                excelapp.Workbooks.Open(path_folder + "\\" + file.Name,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);
                excelapp.Workbooks.Open(path_folder + "\\" + file.Name,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                 Type.Missing, Type.Missing);

                sheet = (Microsoft.Office.Interop.Excel.Worksheet)excelapp.Worksheets[1];
            }

            int count_sh = sheet.UsedRange.Rows.Count + 1;
            List<string> listXLS = new List<string>();
            for (int i = 1; i < count_sh; i++ )
            {
                try
                {
                    Microsoft.Office.Interop.Excel.Range cellRange = sheet.Cells[i, 6];
                    string isbn_xls = null;
                    string price_kazan = null;

                    if (cellRange.Value != null)
                    {
                        isbn_xls = cellRange.Value.ToString();
                        cellRange = sheet.Cells[i, 10];
                        price_kazan = cellRange.Value.ToString();
                    }
                    if (isbn_xls != null)
                    {
                        // чистим isbn
                        StringBuilder sb = new StringBuilder(isbn_xls.Length);
                        foreach (char ch in isbn_xls)
                        {
                            if ((ch >= '0' && ch <= '9'))
                            {
                                sb.Append(ch);
                            }
                        }
                        isbn_xls = sb.ToString();// собранная строка
                        listXLS.Add(isbn_xls + ";" + price_kazan);
                        label1.RefreshData(i.ToString());
                    }
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            ReleaseExcel(excelapp as Object);
            Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (Process p2 in ps2)
            {
                p2.Kill();
            }

            int i_res = 0;
            foreach (string str in richText)
            {
                string[] name_isbn = str.Split(';');
                string isbn_k = name_isbn[0];

                foreach (string str_xls in listXLS)
                {
                    string[] strXlsPrace = str_xls.Split(';');
                    string isbn_xls = strXlsPrace[0];
                    if (isbn_xls == isbn_k)
                    {
                        List<string> res_str = new List<string>();
                        res_str.Add(isbn_k);
                        res_str.Add(name_isbn[1]);
                        res_str.Add(strXlsPrace[1]);
                        res.Add(res_str);
                        i_res++;
                        label2.RefreshData((i_res).ToString());
                    }
                }
            }
            t1.Stop();
            return res;
        }
        private void ReleaseExcel(object excel)
        {
            // Уничтожение объекта Excel.
            Marshal.ReleaseComObject(excel);
            // Вызываем сборщик мусора для немедленной очистки памяти
            GC.GetTotalMemory(true);
        }


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
                label.RefreshData((i_booka+1).ToString());
                string[] name_isbn = richText[i_booka].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];
                
                for (int j = 1; j < 4; j++) // постраничка
                {
                    
                    string site_str = "http://www.booka.ru/search?page=" + j + "&st=title&q=";
                    string strSearch = site_str + HttpUtility.UrlEncode(name); //HttpUtility.UrlEncode(convertedByteString)

                    var page = GetPageCode(strSearch);

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
                                List<string> lst = GetWithIn(txt_isbn, "ISBN:", 28);
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
                                if(isbn == str_isbn)
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
                    var page = GetPageCode(strSearch);
                    
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
                                List<string> lst = GetWithIn(txt_isbn, "ISBN:", 28);
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
                    var page = GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                    if(page != null)
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
                                List<string> lst = GetWithIn(txt_isbn, "ISBN:", 28);
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
                                    var page_price = GetPageCode("http://www.100000-knig.ru" + href);
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

        // goParse_mdk_arbat
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
                    var page = GetPageCode(strSearch);

                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_isbn = new HtmlAgilityPack.HtmlDocument();
                    HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();

                    if (page != null)
                    {
                        doc.LoadHtml(page); // вся страница
                    }
                    else
                    {
                        MessageBox.Show("Сервер долго не отдает страницу!");
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

        // goParse_biblio_globus
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

                string book_id = POST("http://www.bgshop.ru/Default.aspx", "ucSearchBrief$txtSearch=" + isbn);
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
                    var page = GetPageCode(strSearch);

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
                            var page_detail = GetPageCode("http://www.biblio-globus.ru" + link);
                            doc_isbn.LoadHtml(page_detail);
                            HtmlNode NodeISBN = doc_isbn.DocumentNode.SelectSingleNode(".//div[@class='details']");
                            
                            string txt_isbn = NodeISBN.InnerHtml;
                            if (txt_isbn.Contains("ISBN:"))
                            {
                                List<string> lst = GetWithIn(txt_isbn, "ISBN:", 28);
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
                                    HtmlNode NodePrice= doc_isbn.DocumentNode.SelectSingleNode(".//div[@class='details_price']");
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

                var page = GetPageCode(site_str);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 

                if (page != null)
                {
                    doc.LoadHtml(page); // вся страница
                    //link = "http://bookbars.ru";
                    link = "http://bookbars.ru" + GetWithInChString(page, "<h3><a href=\"", '"') + "/";
                    //link += "/";
                }
                else
                {

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

        public MyList<List<string>> goParse_ozon(List<string> richText, MyList<List<string>> res, Label label, Label label1, 
                                                    System.Windows.Forms.Timer t1, 
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

            TorClient Tor = new TorClient();
            Tor.StopTor();
            Tor.StartTor();

            Thread.Sleep(10000);

            for (int i = 0; i < richText.Count(); i++)
            {
                label.RefreshData((i + 1).ToString());
                string[] name_isbn = richText[i].Split(';');
                name = name_isbn[1];
                isbn = name_isbn[0];
                
                if(name.Contains("("))
                {
                    int last_author = name.LastIndexOf('(');
                    name = name.Substring(0, last_author);
                }
                
                string site_str = "http://www.ozon.ru/?context=search&text=" + isbn + "&store=1,0";

                //var page = GetPageCode(site_str, proxi);
                var page = Tor.RunParallel(site_str).Result;

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
                                    List<string> lst = GetWithIn(content_book, "\"price\" content=\"", 8);
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
                                string href_book = GetWithInStr(content_book, "href=\"", "\">");
                                link = "http://www.ozon.ru" + href_book;

                                //var page_book = GetPageCode(link, proxi);
                                var page_book = Tor.RunParallel(link).Result;

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

                        string link_id = GetWithInStr(doc.DocumentNode.InnerHtml, "s.products = \";", ";");
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
            Tor.StopTor();
            return res;
        }

        public MyList<List<string>> goParse_labirint(List<string> richText, MyList<List<string>> res, Label label, Label label1, 
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

            //http://www.labirint.ru/search/9785699793129/?labsearch=1

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

                string site_str = "http://www.labirint.ru/search/"+ isbn + "/?labsearch=1";

                var page = GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                //HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                //HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                if(page != null)
                    doc.LoadHtml(page); // вся страница

                //<meta property="og:url" content="," />
                link = "http://www.labirint.ru/books/";
                link += GetWithInStr(page, "http://img1.labirint.ru/books/", "/big.jpg");
                link += "/";

                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1"); // name 
                if (NodeName != null)
                {
                    if(!NodeName[0].InnerText.Contains("у нас ничего нет"))
                    {
                        string [] arr_name = NodeName[0].InnerText.Split(':');
                        if(arr_name.Count() > 1)
                        {
                            name = arr_name[1].Trim();
                        }
                        else
                        {
                            name = arr_name[0].Trim();
                        }

                        HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//div[@class='publisher']"); // 
                        if (NodePublisher != null)
                        {
                            publisher = NodePublisher[0].InnerText.Split(':')[1].Split(',')[0].Trim();
                            year = NodePublisher[0].InnerText.Split(':')[1].Split(',')[1].Trim();
                        }

                        HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='authors']"); // 
                        if (NodeAuthor != null)
                        {
                            author = NodeAuthor[0].InnerText.Split(':')[1].Trim();
                        }

                        HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//span[@class='buying-pricenew-val-number']"); // price
                        if (NodePrice != null)
                        {
                            price = NodePrice[0].InnerText;
                        }
                        else
                        {
                            HtmlNodeCollection NodePriceOld = doc.DocumentNode.SelectNodes(".//span[@class='buying-price-val-number']"); // price
                            if (NodePriceOld != null)
                            {
                                price = NodePriceOld[0].InnerText;
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
                    else
                    {
                        name = "товар не найден";
                        //string isbn = "";
                        //string link = "";
                        //string publisher = "";
                        //string author = "";
                        //string year = "";
                        //string price = "";
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
            }
            t1.Stop();
            return res;
        }


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

                var page = GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                //HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                //HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                if (page != null)
                    doc.LoadHtml(page); // вся страница

                //<meta property="og:url" content="," />
                link = "http://read.ru/id/";
                link += GetWithInChString(page, "data-book-id=\"", '"');
                link += "/";

                HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1[@class='book_card__header']"); // name 
                if (NodeName != null)
                {
                    name = NodeName[0].InnerText;
                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//div[@class='j-book_autors book_autors']"); // 
                    if (NodeAuthor != null)
                    {
                        if(NodeAuthor[0].InnerText.Contains(":"))
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

        public MyList<List<string>> goParse_kniga(List<string> richText, MyList<List<string>> res, Label label, Label label1,
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

            // http://www.kniga.ru/search/?search_query=9785699782703

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

                string site_str = "http://www.kniga.ru/search/?search_query=" + isbn;

                var page = GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 
                //HtmlAgilityPack.HtmlDocument doc_book = new HtmlAgilityPack.HtmlDocument();
                //HtmlAgilityPack.HtmlDocument doc_price = new HtmlAgilityPack.HtmlDocument();
                if (page != null)
                    doc.LoadHtml(page); // вся страница
                else
                    break;

                //<meta property="og:url" content="," />
                //link = "http://read.ru/id/";
                //link += GetWithInChString(page, "data-book-id=\"", '"');
                //link += "/";
                link = site_str;

                HtmlNodeCollection NodeLink = doc.DocumentNode.SelectNodes(".//div[@class='subNav']/a"); // name
                if (NodeLink != null)
                {
                    // string link = n.GetAttributeValue("href", ""); // получем ссылку
                    link = "http://www.kniga.ru";
                    link += NodeLink[0].GetAttributeValue("href", "");
                    page = GetPageCode(link, proxy);

                    if (page != null)
                        doc.LoadHtml(page); // вся страница

                    HtmlNodeCollection NodeName = doc.DocumentNode.SelectNodes(".//h1"); // 
                    if (NodeName != null)
                    {
                        //if (NodeAuthor[0].InnerText.Contains(":"))
                        name = NodeName[0].InnerText;
                    }

                    HtmlNodeCollection NodeAuthor = doc.DocumentNode.SelectNodes(".//span[@id='authorsList']/a"); // 
                    if (NodeAuthor != null)
                    {
                        author = NodeAuthor[0].InnerText;
                    }

                    HtmlNodeCollection NodePublisher = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[2]/a"); // 
                    if (NodePublisher != null)
                    {
                        publisher = NodePublisher[0].InnerText;
                    }
                    else
                    {
                        HtmlNodeCollection NodePublisher2 = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[1]/a"); // 
                        if (NodePublisher2 != null)
                        {
                            publisher = NodePublisher2[0].InnerText;
                        }
                    }

                    HtmlNodeCollection NodeYear = doc.DocumentNode.SelectNodes(".//*[@id='properties']/p[2]/span[2]"); // 
                    if (NodeYear != null)
                    {
                        year = NodeYear[0].InnerText;
                        StringBuilder sb = new StringBuilder(year.Length);
                        foreach (char ch in year)
                        {
                            if (((ch >= '0' && ch <= '9')))
                            {
                                sb.Append(ch);
                            }
                        }
                        year = sb.ToString();// собранная строка
                    }

                    // .//*[@id='normalPrice']
                    HtmlNodeCollection NodePrice = doc.DocumentNode.SelectNodes(".//*[@id='normalPrice']/span"); // price
                    string str_price = "";
                    if (NodePrice != null)
                    {
                        price = NodePrice[0].InnerText;
                        StringBuilder sb = new StringBuilder(price.Length);

                        foreach (char ch in price)
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
                        NodePrice = doc.DocumentNode.SelectNodes(".//*[@id='normalPrice']"); // price
                        if (NodePrice != null)
                        {
                            str_price = NodePrice[0].InnerText;
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
            t1.Stop();
            return res;
        }

        
        public MyList<List<string>> goParse_biblion(List<string> richText, MyList<List<string>> res, Label label, Label label1,
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

                var page = GetPageCode(site_str, proxy);

                HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument(); // 

                if (page != null)
                    doc.LoadHtml(page); // вся страница

                link = "http://www.biblion.ru/product/";
                link += GetWithInChString(page, "inCart-", '"');
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


        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее заданное количество символов.
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="col_ch"></param>
        /// <returns></returns>
        public List<string> GetWithIn(string str1, string str2, int col_ch)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    if (col_ch > 0)
                    {
                        string str = str1.Substring(i + str2.Length, col_ch);
                        rez.Add(str);
                    }
                    else  // обработка отрицательного значения col_ch при отмотке
                    {
                        i += col_ch;
                        int col = col_ch - col_ch * 2;
                        string str = str1.Substring(i, col);
                        rez.Add(str);
                        i += col;
                    }

                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до заданного (ch).
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public List<string> GetWithInCh(string str1, string str2, char ch)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    //выбор до опред. симвовла
                    //new string(s.TakeWhile(x => x != '{').ToArray()).Length
                    string str11 = str1.Substring(i + str2.Length, str1.Length - (i + str2.Length));
                    string str = new string(str11.TakeWhile(x_i => x_i != ch).ToArray());
                    rez.Add(str);
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до заданного (ch).
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает первую строку
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public string GetWithInChString(string str1, string str2, char ch)
        {
            string rez = "";
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    //выбор до опред. симвовла
                    //new string(s.TakeWhile(x => x != '{').ToArray()).Length
                    string str11 = str1.Substring(i + str2.Length, str1.Length - (i + str2.Length));
                    rez = new string(str11.TakeWhile(x_i => x_i != ch).ToArray());
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит метку str2, и получает после нее символы до str_end.
        /// при отрицательном значении смотрит в обратном направлении )
        /// получает все подстроки после найденных меток
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="ch"></param>
        /// <returns></returns>
        public List<string> GetWithInStrList(string str1, string str2, string str_end)
        {
            List<string> rez = new List<string>();
            int i = 0;  // Числовая переменная, контролирующая итерации цикла
            int x = -1; // Так как метод IndexOf() возвращает "-1" если первое вхождение подстроки не найдено, то приходится использовать вспомагательную, вместо і, что б начать цикл
            //int count = -1; // Записываем количество вхождений (итераций цикла)
            while (i != -1)
            {
                i = str1.IndexOf(str2, x + 1); // получаем индекс первого вхождения  х+1 говорит, что начинать нужно с 0-го индекса
                if (i > -1)
                {
                    int i_str = str1.IndexOf(str_end, i + str2.Length);
                    if (i_str != -1)
                    {
                        int l_str1 = str1.Length;
                        int l_str2 = str2.Length;
                        int str_search = i_str - (i + l_str2);
                        string str = str1.Substring(i + l_str2, str_search);
                        rez.Add(str);
                    }
                }
                x = i; // присваиваем номер индекса первого значения, что б потом (х+1) начать со следующего
                // count++;  // Увеличиваем на единицу наше количество
            }
            return rez;
        }

        /// <summary>
        /// В строке str1 находит первое вхождение метки str2, и получает после нее символы до метки str_end.
        /// </summary>
        /// <param name="str1"></param>
        /// <param name="str2"></param>
        /// <param name="str_end"></param>
        /// <returns></returns>
        public string GetWithInStr(string str1, string str2, string str_end)
        {
            string rez = "";
            int i = 0;
            i = str1.IndexOf(str2, 0); // получаем индекс первого вхождения с 0-го индекса
            if (i > -1)
            {
                int i_str2 = str1.IndexOf(str_end, i + str2.Length);
                if (i_str2 != -1)
                {
                    int l_str1 = str1.Length;
                    int l_str2 = str2.Length;
                    int str_search = i_str2 - (i + l_str2);
                    rez = str1.Substring(i + l_str2, str_search);
                }
            }
            return rez;
        }
    }
    
}
