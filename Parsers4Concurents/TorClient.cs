using com.LandonKey.SocksWebProxy;
using com.LandonKey.SocksWebProxy.Proxy;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace Parsers4Сompetitor
{
    class TorClient
    {
        static string pathTOR = Directory.GetCurrentDirectory() + "\\Browser\\firefox.exe";

        public async Task<string> RunParallel(string url)
        {
            var locker = new object();
            var proxy = new SocksWebProxy(new ProxyConfig(
                //This is an internal http->socks proxy that runs in process
                IPAddress.Parse("127.0.0.1"),
                //This is the port your in process http->socks proxy will run on
                12345,
                //This could be an address to a local socks proxy (ex: Tor / Tor Browser, If Tor is running it will be on 127.0.0.1)
                IPAddress.Parse("127.0.0.1"),
                //This is the port that the socks proxy lives on (ex: Tor / Tor Browser, Tor is 9150)
                9150,
                //This Can be Socks4 or Socks5
                ProxyConfig.SocksVersion.Five
                ));

            WebClient client = new WebClient();
            //client.Proxy = proxy.IsActive() ? proxy : null;
            client.Proxy = proxy;
            var doc = new HtmlAgilityPack.HtmlDocument();
            Uri siteUri = new Uri(url);

            string html = await client.DownloadStringTaskAsync(siteUri);
            return html;
        }

        public void StartTor()
        {
            ProcessStartInfo si = new ProcessStartInfo(pathTOR);
            si.WindowStyle = ProcessWindowStyle.Hidden;
            Process.Start(si);
        }

        public void StopTor()
        {
            //string name_proc = "tor";//процесс, который нужно убить
            System.Diagnostics.Process[] etc = System.Diagnostics.Process.GetProcesses();//получим процессы
            foreach (System.Diagnostics.Process nameProc in etc)//обойдем каждый процесс
            {
                if (nameProc.ProcessName == "firefox") nameProc.Kill();//найдем нужный и убьем
                if (nameProc.ProcessName == "tor") nameProc.Kill();//найдем нужный и убьем
                //richTextBox1.Text += nameProc.ProcessName.ToLower() + "\n";
            }
        }
    }
}
