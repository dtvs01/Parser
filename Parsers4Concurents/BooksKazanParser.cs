using HtmlAgilityPack;
using Parsers4Concurents;
using SharpCompress.Common;
using SharpCompress.Reader;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Parsers4Сompetitor
{
    class BooksKazanParser
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

        public MyList<List<string>> goParse_bookskazan(List<string> richText, MyList<List<string>> res, Label label1, Label label2, System.Windows.Forms.Timer t1)
        {

            var page = Responses.GetPageCode("http://www.bookskazan.ru/article/452/10/");

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
                //MessageBox.Show(ex.Message);
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
            for (int i = 1; i < count_sh; i++)
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
                    //MessageBox.Show(e.Message);
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
    }
}
