using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Parsers4Сompetitor
{
    static class Responses
    {
        public static string GetPageCode(string strURL, string proxyAddress = "", string username = "", string password = "")
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

                if (proxyAddress != "" && proxyAddress != null)
                {
                    WebProxy myProxy = new WebProxy(proxyAddress, true);
                    //Uri newUri = new Uri(proxyAddress);
                    //myProxy = (WebProxy)request.Proxy;
                    //myProxy.Address = newUri;

                    if (username != "" && password != "")
                    {
                        myProxy.Credentials = new NetworkCredential(username, password);
                    }
                    request.Proxy = myProxy;
                }


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
                //else
                //{
                //    PageHTML = null;
                //    return PageHTML;
                //}

                //Реализуем считывание символов из потока байтов в определенной кодировке.
                //GetResponseStream -  возвращает поток данных из  интернет-ресурса
                //using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.Default))
            }
            catch (Exception ex)
            {
                PageHTML = null;
                return PageHTML;
            }
            return PageHTML;
        }

        public static string GetPageCodeBG(string strURL)
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
                //else
                //{
                    //MessageBox.Show("response == null!");
                    //return PageHTML;
                //}

                //Реализуем считывание символов из потока байтов в определенной кодировке.
                //GetResponseStream -  возвращает поток данных из  интернет-ресурса
                //using (StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.Default))


            }
            catch (Exception ex)
            {
                PageHTML = null;
                return PageHTML;
            }
            return PageHTML;
        }


        /// <summary>
        /// Отсылка POST-запроса на URL 
        /// </summary>
        /// <param name="Url"></param>
        /// <param name="Data"></param>
        /// <returns></returns>
        public static string POST(string Url, string Data)
        {
            string Out = null;
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
                return Out;
            }
            return Out;
        }
    }
}
