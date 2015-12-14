using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Parsers4Сompetitor;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace Parsers4Concurents
{
    public partial class Form1 : Form
    {
        // Формируем строку с параметрами подключения к файлу базы данных
        // Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\bin\Debug\labirint.accdb
        static string path_folder = Directory.GetCurrentDirectory();
        string connectionString =
        "Provider=Microsoft.ACE.OLEDB.12.0;" +
        @"Data Source=" + path_folder + "\\labirint_3.accdb";
        public Form1()
        {
            InitializeComponent();
            ListOMSK.OnAdd += new EventHandler(ListOMSK_OnAdd);
            ListKazan.OnAdd += new EventHandler(ListKazan_OnAdd);
            ListBooka.OnAdd += new EventHandler(ListBooka_OnAdd);
            ListProdalit.OnAdd += new EventHandler(ListProdalit_OnAdd);
            List100000knig.OnAdd += new EventHandler(List100000knig_OnAdd);
            List_mdk_arbat.OnAdd += new EventHandler(List_mdk_arbat_OnAdd);
            List_biblio.OnAdd += new EventHandler(List_biblio_OnAdd);
            List_bookbars.OnAdd += new EventHandler(List_bookbars_OnAdd);
            List_ozon.OnAdd += new EventHandler(List_ozon_OnAdd);
            List_labirint.OnAdd += new EventHandler(List_labirint_OnAdd);
            List_read.OnAdd += new EventHandler(List_read_OnAdd);
            List_kniga.OnAdd += new EventHandler(List_kniga_OnAdd);
            List_biblion.OnAdd += new EventHandler(List_biblion_OnAdd);

            all_elements1.OnAdd += new EventHandler(all_elements1_OnAdd);
            all_elements2.OnAdd += new EventHandler(all_elements2_OnAdd);
            all_elements3.OnAdd += new EventHandler(all_elements3_OnAdd);
            all_elements4.OnAdd += new EventHandler(all_elements4_OnAdd);
            all_elements5.OnAdd += new EventHandler(all_elements5_OnAdd);
            all_elements6.OnAdd += new EventHandler(all_elements6_OnAdd);
            all_elements7.OnAdd += new EventHandler(all_elements7_OnAdd);
            all_elements8.OnAdd += new EventHandler(all_elements8_OnAdd);
            all_elements9.OnAdd += new EventHandler(all_elements9_OnAdd);
            all_elements10.OnAdd += new EventHandler(all_elements10_OnAdd);
        }

        static MyList<List<string>> ListOMSK = new MyList<List<string>>();
        static MyList<List<string>> ListKazan = new MyList<List<string>>(); //
        static MyList<List<string>> ListBooka = new MyList<List<string>>();
        static MyList<List<string>> ListProdalit = new MyList<List<string>>();
        static MyList<List<string>> List100000knig = new MyList<List<string>>();
        static MyList<List<string>> List_mdk_arbat = new MyList<List<string>>(); // 
        static MyList<List<string>> List_biblio = new MyList<List<string>>();
        static MyList<List<string>> List_bookbars = new MyList<List<string>>();
        static MyList<List<string>> List_ozon = new MyList<List<string>>();
        static MyList<List<string>> List_labirint = new MyList<List<string>>();
        static MyList<List<string>> List_read = new MyList<List<string>>();
        static MyList<List<string>> List_kniga = new MyList<List<string>>();
        static MyList<List<string>> List_biblion = new MyList<List<string>>();

        MyList<List<string>> all_elements1 = new MyList<List<string>>();
        MyList<List<string>> all_elements2 = new MyList<List<string>>();
        MyList<List<string>> all_elements3 = new MyList<List<string>>();
        MyList<List<string>> all_elements4 = new MyList<List<string>>();
        MyList<List<string>> all_elements5 = new MyList<List<string>>();
        MyList<List<string>> all_elements6 = new MyList<List<string>>();
        MyList<List<string>> all_elements7 = new MyList<List<string>>();
        MyList<List<string>> all_elements8 = new MyList<List<string>>();
        MyList<List<string>> all_elements9 = new MyList<List<string>>();
        MyList<List<string>> all_elements10 = new MyList<List<string>>();

        System.Data.DataTable reportDetailTable;
        System.Data.DataTable labirintDetailTable;
        System.Data.DataTable omskDetailTable;
        System.Data.DataTable kazanDetailTable;
        System.Data.DataTable bookaDetailTable;
        System.Data.DataTable prodalitDetailTable;
        System.Data.DataTable knig_100000DetailTable;
        System.Data.DataTable mdk_arbatDetailTable; // 
        System.Data.DataTable biblio_globus_DetailTable;
        System.Data.DataTable bookbars_DetailTable;
        System.Data.DataTable ozon_DetailTable;
        System.Data.DataTable labirint_DetailTable;
        System.Data.DataTable read_DetailTable;
        System.Data.DataTable kniga_DetailTable;
        System.Data.DataTable biblion_DetailTable;

        public void all_elements1_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements1, "labirint", dataGridView3, labirintDetailTable, "1");
        }

        public void all_elements2_OnAdd(object sender, EventArgs e)
        {
           // ParseALLWrite(all_elements2, "labirint", dataGridView3, labirintDetailTable, "2");
        }

        public void all_elements3_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements3, "labirint", dataGridView3, labirintDetailTable, "3");
        }

        public void all_elements4_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements4, "labirint", dataGridView3, labirintDetailTable, "4");
        }

        public void all_elements5_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements5, "labirint", dataGridView3, labirintDetailTable, "5");
        }

        public void all_elements6_OnAdd(object sender, EventArgs e)
        {
           // ParseALLWrite(all_elements6, "labirint", dataGridView3, labirintDetailTable, "6");
        }

        public void all_elements7_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements7, "labirint", dataGridView3, labirintDetailTable, "7");
        }

        public void all_elements8_OnAdd(object sender, EventArgs e)
        {
           // ParseALLWrite(all_elements8, "labirint", dataGridView3, labirintDetailTable, "8");
        }

        public void all_elements9_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements9, "labirint", dataGridView3, labirintDetailTable, "9");
        }

        public void all_elements10_OnAdd(object sender, EventArgs e)
        {
            //ParseALLWrite(all_elements10, "labirint", dataGridView3, labirintDetailTable, "10");
        }

        public void ListOMSK_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(ListOMSK, "omskbook", dataGridView1, omskDetailTable, null, "omsk");
        }

        public void ListKazan_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(ListKazan, "book_kazan", dataGridView2, kazanDetailTable);
        }

        public void ListBooka_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(ListBooka, "booka", dataGridView4, bookaDetailTable);
        }

        public void ListProdalit_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(ListProdalit, "prodalit", dataGridView5, prodalitDetailTable);
        }

        public void List100000knig_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List100000knig, "100000knig", dataGridView6, knig_100000DetailTable);
        }

        public void List_mdk_arbat_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_mdk_arbat, "mdk_arbat", dataGridView3, mdk_arbatDetailTable);
        }

        public void List_biblio_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_biblio, "biblio_globus", dataGridView7, biblio_globus_DetailTable);
        }

        public void List_bookbars_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_bookbars, "bookbars", dataGridView8, bookbars_DetailTable);
        }

        public void List_ozon_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_ozon, "ozon", dataGridView9, ozon_DetailTable);
        }

        public void List_labirint_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_labirint, "labirint", dataGridView10, labirint_DetailTable);
        }

        public void List_read_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_read, "read", dataGridView11, read_DetailTable);
        }

        public void List_kniga_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_kniga, "kniga", dataGridView12, kniga_DetailTable);
        }

        public void List_biblion_OnAdd(object sender, EventArgs e)
        {
            ParseALLWrite(List_biblion, "biblion", dataGridView13, biblion_DetailTable);
        }

        private void SetColumnsDGV(DataGridView dataGridView)
        {
            dataGridView.Columns.Add("link", "ссылка");
            dataGridView.Columns.Add("isbn", "ISBN");
            dataGridView.Columns.Add("book-name", "Название");
            dataGridView.Columns.Add("author", "Автор");
            dataGridView.Columns.Add("price", "Цена");
            dataGridView.Columns.Add("publisher", "Издательство");
        }

        int num_parse = 1;
        int num_parse_osmo = 1;
        private void ParseALLWrite(MyList<List<string>> Lists, string el_name, DataGridView dataGridView, System.Data.DataTable Table, string num = "", string parse_name = "")
        {
            try
            {
                if (Lists.Count > 0)
                {
                    List<string> lst = Lists.Last();
                    if (lst.Count < 4 && lst.Count > 0)
                    {
                        // блок для парсеров по заявленному csv
                        string str = lst[0] + ";" + lst[1] + ";" + lst[2] + ";";
                        System.IO.StreamWriter writer = new System.IO.StreamWriter("parse_" + parse_name + "_" + el_name + ".csv", true, System.Text.Encoding.GetEncoding(1251));
                        writer.WriteLine(str);
                        writer.Close();

                        
                        int i_row = Table.Rows.Count + 1;
                        string[] row = new string[lst.Count];

                        dataGridView.Invoke(new MethodInvoker(delegate()
                        {
                            dataGridView.Rows.Add("", lst[0], lst[1], "", lst[2]);
                            //row[i] = lst[i];
                            
                        }));
                        //Table.Rows.Add(row);
                        //dataGridView.SetTableIn(Table);
                         
                        num_parse_osmo++;
                    }
                    else
                    {
                        string link = null;
                        string isbn = null;
                        string name = null;
                        string author = null;
                        string price = null;
                        string izdat = null;
                        string dosup = null;
                        string year = null;

                        if (lst.Count > 0)
                        {
                            link   = lst[0];
                            isbn   = lst[1];
                            name   = lst[2];
                            author = lst[3];
                            price  = lst[4];
                            izdat  = lst[5];
                            year   = lst[6];

                            dataGridView.Invoke(new MethodInvoker(delegate ()
                            {
                                dataGridView.Rows.Add(link, isbn, name, author, price, izdat, year);
                            }));

                            if (izdat.Contains("Издательство: "))
                            {
                                izdat = izdat.Replace("Издательство: ", "");
                                if (izdat.Contains("г."))
                                {
                                    string[] izd = izdat.Split(',');
                                    izdat = izd[0];
                                    // преобразуем цену
                                    StringBuilder sb = new StringBuilder(price.Length);
                                    foreach (char ch in izd[1])
                                    {
                                        if (ch >= '0' && ch <= '9')
                                        {
                                            sb.Append(ch);
                                        }
                                    }
                                    year = sb.ToString();// собранная строка

                                }

                            }
                            if (isbn.Contains("&nbsp; скрыть"))
                            {
                                string[] isbn_all = isbn.Split(',');
                                foreach (string isbn_str in isbn_all)
                                {
                                    string isbn_plus = null;
                                    if (isbn_str.Contains("&nbsp; все"))
                                    {
                                        isbn_plus = isbn_str.Substring(0, isbn_str.Length - 10);
                                    }
                                    else if (isbn_str.Contains("&nbsp; скрыть"))
                                    {
                                        isbn_plus = isbn_str.Substring(0, isbn_str.Length - 13);
                                    }
                                    else
                                    {
                                        isbn_plus = isbn_str;
                                    }

                                    if (isbn_plus.Contains("ISBN: "))
                                    {
                                        isbn_plus = isbn_plus.Substring(6, isbn_plus.Length - 6);
                                    }

                                    //Table.Rows.Add(new Object[] { num_parse.ToString(), link, isbn_plus.TrimStart(' '), name, price, izdat, dosup });
                                    //dataGridView3.SetTableIn(Table);

                                    if (name.Contains('"'))
                                    {
                                        name = replace_quotes(name);
                                    }
                                    if (izdat.Contains('"'))
                                    {
                                        izdat = replace_quotes(izdat);
                                    }

                                    isbn_plus = isbn_plus.TrimStart(' ');
                                    if (isbn_plus.Contains('"'))
                                    {
                                        isbn_plus = isbn_plus.Replace("\"", "");
                                    }

                                    //label11.RefreshData(num_parse.ToString());
                                    /*
                                    try
                                    {
                                        OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
                                        // создаем объект OleDbCommand
                                        OleDbCommand myOleDbCommand = myOleDbConnection.CreateCommand();

                                        string sSQL = @" INSERT INTO [labirint] ([link], [isbn], [name_book], [publishing], [year], [price], [available]) VALUES (";
                                        sSQL += '"' + link + '"' + ',';
                                        sSQL += '"' + isbn_plus + '"' + ',';
                                        sSQL += '"' + name + '"' + ',';
                                        sSQL += '"' + izdat + '"' + ','; // year
                                        sSQL += '"' + year + '"' + ',';
                                        sSQL += '"' + price + '"' + ',';
                                        sSQL += '"' + dosup + '"' + ')' + ';';

                                        myOleDbCommand.CommandText = sSQL;

                                        // открываем соединение с БД с помощью метода Open() объекта OleDbConnection
                                        // открываем соединение с БД с помощью метода Open() объекта OleDbConnection
                                        if (myOleDbConnection.State != ConnectionState.Open)
                                        {
                                            myOleDbConnection.Open();
                                        }
                                        //myOleDbConnection.Open();
                                        myOleDbCommand.ExecuteNonQuery();// выполняем запрос
                                        myOleDbConnection.Close();       // закрываем базу
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }                                   
                                    */
                                    string str_lab = /*num_parse.ToString() + ";" + */link + ";" + isbn_plus + ";" + name + ";" + price + ";" + izdat + ";" + year + ";" + dosup + ";";
                                    try
                                    {
                                        using (System.IO.StreamWriter writer_lab = new System.IO.StreamWriter("parse_file_" + el_name + "_" + num + ".csv", true, System.Text.Encoding.GetEncoding(1251)))
                                        {
                                            writer_lab.WriteLine(str_lab);
                                            //writer_lab.Close();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        MessageBox.Show(e.Message);
                                    }
                                    num_parse++;
                                    Lists.Clear();
                                }
                            }
                            else
                            {
                                if (isbn.Contains("ISBN: "))
                                {
                                    isbn = isbn.Substring(6, isbn.Length - 6);
                                }
                                //Table.Rows.Add(new Object[] { num_parse.ToString(), link, isbn, name, price, izdat, dosup });
                                //dataGridView3.SetTableIn(Table);

                                //label11.RefreshData(num_parse.ToString());

                                if (name.Contains('"'))
                                {
                                    name = replace_quotes(name);
                                }

                                if (izdat.Contains('"'))
                                {
                                    izdat = replace_quotes(izdat);
                                }

                                if (isbn.Contains('"'))
                                {
                                    isbn = isbn.Replace("\"", "");
                                }

                                /*
                                try
                                {
                                    OleDbConnection myOleDbConnection = new OleDbConnection(connectionString);
                                    // создаем объект OleDbCommand
                                    OleDbCommand myOleDbCommand = myOleDbConnection.CreateCommand();

                                    string sSQL = @" INSERT INTO [labirint] ([link], [isbn], [name_book], [publishing], [year], [price], [available]) VALUES (";
                                    sSQL += '"' + link + '"' + ',';
                                    sSQL += '"' + isbn + '"' + ',';
                                    sSQL += '"' + name + '"' + ',';
                                    sSQL += '"' + izdat + '"' + ',';
                                    sSQL += '"' + year + '"' + ',';
                                    sSQL += '"' + price + '"' + ',';
                                    sSQL += '"' + dosup + '"' + ')' + ';';

                                    myOleDbCommand.CommandText = sSQL;

                                    // открываем соединение с БД с помощью метода Open() объекта OleDbConnection
                                    // открываем соединение с БД с помощью метода Open() объекта OleDbConnection
                                    if (myOleDbConnection.State != ConnectionState.Open)
                                    {
                                        myOleDbConnection.Open();
                                    }
                                    //myOleDbConnection.Open();
                                    myOleDbCommand.ExecuteNonQuery();// выполняем запрос
                                    myOleDbConnection.Close();       // закрываем базу
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show(e.Message);
                                }
                                */

                                string str_lab = /*num_parse.ToString() + ";" + */link + ";" + isbn + ";" + name + ";" + author + ";" + price + ";" + izdat + ";" + year + ";"; // + dosup + ";";
                                try
                                {
                                    using (System.IO.StreamWriter writer_lab = new System.IO.StreamWriter("parse_file_" + el_name + "_" + num + ".csv", true, System.Text.Encoding.GetEncoding(1251)))
                                    {
                                        writer_lab.WriteLine(str_lab);
                                        //writer_lab.Close();
                                    }                                    
                                }
                                catch (Exception e)
                                {
                                    MessageBox.Show(e.Message);
                                }
                                num_parse++;
                                Lists.Clear();
                            }
                        }

                    }

                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            
            
        }

        private string replace_quotes(string m)
        {
            int at = 0;
            int pos = 0;
            while ((pos < m.Length) && (at > -1))
            {
                at = m.IndexOf('"', pos);
                if (at == -1) break; // Выход из цикла если больше нет нужного символа в строке   
                pos = at + 1;

                Regex R = new Regex(@"[^wd]");

                // Если конец строки заканчивается кавычкой
                if (pos == m.Length) { m = R.Replace(m, "'", 1, at); break; }

                // Если впереди кавычки идут буквы и цифры, - ставим открывающую "«"
                // Если любые другие символы, - ставим закрывающую "»"
                Match match = R.Match(m, pos, 1);
                if (match.Success) { m = R.Replace(m, "'", 1, at); }
                else { m = R.Replace(m, "'", 1, at); }
            }

            return m;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "csv файлы (*.csv)|*.csv|txt файлы(*.txt)|*.txt";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                System.IO.StreamReader streamReader;
                streamReader = new System.IO.StreamReader(openFileDialog1.FileName);
                richTextBox1.Text = streamReader.ReadToEnd();
                streamReader.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OmskParser OmskParse = new OmskParser();

            pictureBox1.Visible = false;
            pictureBox2.Visible = true;
            button2.Enabled = false;
            button4.Enabled = false;
            omskDetailTable = new System.Data.DataTable("omskParse");

            // Define all the columns once.
            DataColumn[] cols_omsk ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            omskDetailTable.Columns.AddRange(cols_omsk);

            omskDetailTable.Rows.Clear();



            if(dataGridView1.Rows.Count > 0)
            {
                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
            }
            num_parse_osmo = 1;
            t = 0;

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*omsk*"))
            {
                file.Delete();
            }

            timer1.Start();
            timer2.Start();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> task_omsk = new Task<MyList<List<string>>>(() => OmskParse.goParse_omskbook(richText, ListOMSK, label4, label1, timer1)); // 0, 35157,
                task_omsk.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button2.Enabled = true;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            SetColumnsDGV(dataGridView1);
            SetColumnsDGV(dataGridView2);
            SetColumnsDGV(dataGridView4);
            SetColumnsDGV(dataGridView5);
            SetColumnsDGV(dataGridView6);
            SetColumnsDGV(dataGridView3);
            SetColumnsDGV(dataGridView7);
            SetColumnsDGV(dataGridView8);
            SetColumnsDGV(dataGridView9);
            SetColumnsDGV(dataGridView10);
            SetColumnsDGV(dataGridView11);
            SetColumnsDGV(dataGridView12);
            SetColumnsDGV(dataGridView13);

            reportDetailTable = new System.Data.DataTable("ReportParse");

            // Define all the columns once.
            DataColumn[] cols ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            reportDetailTable.Columns.AddRange(cols);

            labirintDetailTable = new System.Data.DataTable("labirint");

            // Define all the columns once.
            DataColumn[] cols_lab ={
                                  new DataColumn("id",typeof(String)),
                                  new DataColumn("link",typeof(String)),
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String)),
                                  new DataColumn("Издательство",typeof(String)),
                                  new DataColumn("Доступность",typeof(String)),
                              };

            labirintDetailTable.Columns.AddRange(cols_lab);
        }

        
        private void timer1_Tick(object sender, EventArgs e)
        {
            t++;
            label5.Text = t.ToString();
        }

        int t = 0;
        private void timer2_Tick(object sender, EventArgs e)
        {
            if (timer1.Enabled)
            {
                //dataGridView1.Refresh();
            }
            else 
            {
                dataGridView1.Refresh();
                timer2.Stop();
                pictureBox1.Visible = true;
                pictureBox2.Visible = false;
                button3.Enabled = true;
                button4.Enabled = true;
                button2.Enabled = false;
            }
        }

        private void timer_lab_Tick(object sender, EventArgs e)
        {
            //dataGridView3.Refresh();
        }

        public void ExportToExcel(DataGridView grid, string name_parse)
        {
            //ApplicationClass Excel = new ApplicationClass();

            string path_folder = Directory.GetCurrentDirectory();

            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();

            XlReferenceStyle RefStyle = Excel.ReferenceStyle;
            Excel.Visible = true;
            Workbook wb = null;
            String TemplatePath = path_folder + @"\Parse.xltx";
            try
            {
                //app.Workbooks.Open(fileex);
                wb = Excel.Workbooks.Open(TemplatePath); // !!! 
            }
            catch (System.Exception ex)
            {
                throw new Exception("Не удалось загрузить шаблон для экспорта " + TemplatePath + "\n" + ex.Message);
            }
            Worksheet ws = wb.Worksheets.get_Item(1) as Worksheet;

            //int iRows = ws.UsedRange.Rows.Count; // последняя заполненная строка

            for (int j = 0; j < grid.Columns.Count; ++j)
            {
                (ws.Cells[1, j + 1] as Range).Value2 = grid.Columns[j].HeaderText;
                for (int i = 0; i < grid.Rows.Count; ++i)
                {
                    object Val = grid.Rows[i].Cells[j].Value;
                    if (Val != null)
                        (ws.Cells[i + 2, j + 1] as Range).Value2 = Val.ToString();
                }
            }

            ws.Columns.EntireColumn.AutoFit();
            Excel.ReferenceStyle = RefStyle;

            // записываем в файл .xls с идентификатором по дате
            // и убиваем процесс Excel
            //string serverNum = server[3];
            DateTime dateNow = DateTime.Now;
            string time = dateNow.ToLocalTime().ToString();
            string[] date = time.Split(' ');
            string str_time = date[1].Replace(':', '_');
            Excel.DisplayAlerts = false;
            //wb.Save();
            wb.SaveAs(path_folder + @"\Parse_" + name_parse + "_" + date[0] + " " + str_time + ".xls");
            wb.Close(false, false, false);
            ReleaseExcel(Excel as Object);
            Process[] ps2 = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (Process p2 in ps2)
            {
                p2.Kill();
            }
        }

        private void ReleaseExcel(object excel)
        {
            // Уничтожение объекта Excel.
            Marshal.ReleaseComObject(excel);
            // Вызываем сборщик мусора для немедленной очистки памяти
            GC.GetTotalMemory(true);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView1, "omsk");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            button2.Enabled = true;
            button3.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Request request = new Request();
            for (int i = 1; i <= 10; i++)
            {
                request.LoadXML(@"E:\catalog" + i + ".xml", "http://www.labirint.ru/smcatalog" + i + ".xml");
            }
        }

        //delegate void GetBook(MyList<List<string>> ElementsForGrid, string file);
        Task task1;
        private void button8_Click(object sender, EventArgs e)
        {
            timer_lab.Start();
            timerLabTime.Start();
            timerAllBase.Start();

            // установка файлов потокам
            string path_folder = Directory.GetCurrentDirectory();
            string file1 = path_folder + @"\catalog_labirint\txt_file_0.txt";
            string file2 = path_folder + @"\catalog_labirint\txt_file_1.txt";
            string file3 = path_folder + @"\catalog_labirint\txt_file_2.txt";
            string file4 = path_folder + @"\catalog_labirint\txt_file_3.txt";
            string file5 = path_folder + @"\catalog_labirint\txt_file_4.txt";
            string file6 = path_folder + @"\catalog_labirint\txt_file_5.txt";
            string file7 = path_folder + @"\catalog_labirint\txt_file_6.txt";
            string file8 = path_folder + @"\catalog_labirint\txt_file_7.txt";
            string file9 = path_folder + @"\catalog_labirint\txt_file_8.txt";
            string file10 = path_folder + @"\catalog_labirint\txt_file_9.txt";

            // проверка точек останова
            string[] lines1 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_1.csv");
            int f_last_1 = 0;
            if (lines1.Count() != 0)
            {
                f_last_1 = lines1.Count();
            }
            
            string[] lines2 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_2.csv");
            int f_last_2 = 0;
            if (lines2.Count() != 0)
            {
                f_last_2 = lines2.Count();
            }

            string[] lines3 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_3.csv");
            int f_last_3 = 0;
            if (lines1.Count() != 0)
            {
                f_last_3 = lines3.Count();
            }

            string[] lines4 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_4.csv");
            int f_last_4 = 0;
            if (lines1.Count() != 0)
            {
                f_last_4 = lines4.Count();
            }

            string[] lines5 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_5.csv");
            int f_last_5 = 0;
            if (lines5.Count() != 0)
            {
                f_last_5 = lines5.Count();
            }

            string[] lines6 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_6.csv");
            int f_last_6 = 0;
            if (lines6.Count() != 0)
            {
                f_last_6 = lines6.Count();
            }

            string[] lines7 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_7.csv");
            int f_last_7 = 0;
            if (lines7.Count() != 0)
            {
                f_last_7 = lines7.Count();
            }

            string[] lines8 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_8.csv");
            int f_last_8 = 0;
            if (lines1.Count() != 0)
            {
                f_last_8 = lines8.Count();
            }

            string[] lines9 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_9.csv");
            int f_last_9 = 0;
            if (lines1.Count() != 0)
            {
                f_last_9 = lines9.Count();
            }

            string[] lines10 = System.IO.File.ReadAllLines(path_folder + @"\parse_file_labirint_10.csv");
            int f_last_10 = 0;
            if (lines10.Count() != 0)
            {
                f_last_10 = lines10.Count();
            }

            // запуск потоков
            //var task1 = Task.Factory.StartNew(GetListTxt(all_elements1, file1, f_last_1, label29, label30));
            /*
            task1 = new Task(() => GetListTxt(all_elements1, file1, f_last_1, label29, label30)); //
            task1.Start();

            Task task2 = new Task(() => GetListTxt(all_elements2, file2, f_last_2, label32, label31)); //
            task2.Start();

            Task task3 = new Task(() => GetListTxt(all_elements3, file3, f_last_3, label34, label33)); //
            task3.Start();

            Task task4 = new Task(() => GetListTxt(all_elements4, file4, f_last_4, label36, label35)); //
            task4.Start();

            Task task5 = new Task(() => GetListTxt(all_elements5, file5, f_last_5, label38, label37)); //
            task5.Start();

            Task task6 = new Task(() => GetListTxt(all_elements6, file6, f_last_6, label40, label39)); //
            task6.Start();

            Task task7 = new Task(() => GetListTxt(all_elements7, file7, f_last_7, label42, label41)); //
            task7.Start();

            Task task8 = new Task(() => GetListTxt(all_elements8, file8, f_last_8, label44, label43)); //
            task8.Start();

            Task task9 = new Task(() => GetListTxt(all_elements9, file9, f_last_9, label46, label45)); //
            task9.Start();

            Task task10 = new Task(() => GetListTxt(all_elements10, file10, f_last_10, label48, label47)); //
            task10.Start();
             * */
        }

        /*
        // парсинг по xml
        public MyList<List<string>> GetList(MyList<List<string>> all_elements, string file, int f_last, string proxi = "")
        {
            Request request = new Request();
            List<string> element = new List<string>();
            //List<List<string>> all_elements = new List<List<string>>();
            WebClient client = new WebClient();
            if (proxi != "")
            {
                client.Proxy = new WebProxy(proxi);
            }
            
            //client = request.WebClientProxi("95.85.12.187:3128");
            int it = 0;
            using (XmlReader xml = XmlReader.Create(file))
            {
                while (xml.Read())
                {
                    switch (xml.NodeType)
                    {
                        case XmlNodeType.Element:
                            // нашли элемент 
                            if (xml.Name == "loc")
                            {
                                string link = xml.ReadElementContentAsString();
                                if (link.Contains("/books/"))
                                {
                                    element = request.GetPageAgility(link, client);
                                    all_elements.Add(element);
                                    it++;
                                }
                            }
                         break;
                    }

                }
            }
            return all_elements;
        }
        */

        public void AllLinkCounter()
        {
            /*
            int allLinks = int.Parse(label29.Text)+
                           int.Parse(label32.Text)+
                           int.Parse(label34.Text)+ 
                           int.Parse(label36.Text)+
                           int.Parse(label38.Text)+
                           int.Parse(label40.Text)+
                           int.Parse(label42.Text)+
                           int.Parse(label44.Text)+
                           int.Parse(label46.Text)+
                           int.Parse(label48.Text);
            label53.Text = allLinks.ToString();
             */
        }

        // парсинг по txt
        /*
        public void GetListTxt(MyList<List<string>> all_elements, string file, int f_last, System.Windows.Forms.Label task_count = null, System.Windows.Forms.Label task_parse = null)
        {
            Request request = new Request();
            List<string> element = new List<string>();
            WebClient client = new WebClient();

            string[] lines = System.IO.File.ReadAllLines(file);
            int linesCount = lines.Count();
            if (task_count != null)
            {
                task_count.RefreshData(linesCount.ToString());
            }

            int i = 0;
            for (i += f_last; i < linesCount; i++)
            {
                element = request.GetPageAgility(lines[i], client);
                all_elements.Add(element);

                if (task_parse != null)
                {
                    task_parse.RefreshData(i.ToString());
                }
            }
        }
        */

        int t_lab = 0;
        private void timerLabTime_Tick(object sender, EventArgs e)
        {
            t_lab++;
            //label13.Text = t_lab.ToString();
        }

        /*
        private void button5_Click(object sender, EventArgs e)
        {
            WriteCSVToExcel(System.Windows.Forms.Application.StartupPath,
                System.Windows.Forms.Application.StartupPath + "\\temp.xlsx");
            //label15.Text = strXLSFile;
        }
        */

        string strXLSFile = null;
        private void WriteCSVToExcel(string folder, string xlsx)
        {
            var lst = new List<String>();
            DirectoryInfo dir = new DirectoryInfo(folder);
            foreach (FileInfo file in dir.GetFiles("*parse_file*"))
            {
                lst.Add(file.Name);
            }

            List<string[]> res_file = new List<string[]>();

            foreach (string file_name in lst)
            {
                res_file.Add(File.ReadAllLines(file_name, System.Text.Encoding.GetEncoding(1251)));
            }

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = app.Workbooks.Open(xlsx, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet sheet = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            foreach (string[] arLines in res_file)
            {
                int iRows = sheet.UsedRange.Rows.Count + 1;
                for (int j = 0; j < arLines.Count(); j++)
                {
                    var str = arLines[j].Split(';');
                    for (int i = 0; i < str.Count(); i++)
                    {
                        sheet.Cells[iRows + j, i + 1] = str[i];  //добавление в ячейку значений, разделённых разделителем
                    }
                }
            }

            DateTime dateNow = DateTime.Now;
            string time = dateNow.ToLocalTime().ToString();
            string[] date = time.Split(' ');
            string str_time = date[1].Replace(':', '_');
            strXLSFile = System.Windows.Forms.Application.StartupPath + "\\parse_labirint" + "_" + date[0] + "_" + str_time + ".xlsx";
            wb.SaveAs(strXLSFile, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wb.Close();
            app.Quit();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            var lst = new List<String>();
            DirectoryInfo dir = new DirectoryInfo(@"E:\Работа\май 2015\26_05_15\catalog_labirint");
            foreach (FileInfo file in dir.GetFiles("*smcatalog*"))
            {
                lst.Add(file.Name);
            }

            List<List<string>> AllTXT = new List<List<string>>();
            
            foreach (string file_name in lst)
            {
                List<string> list_link = new List<string>();
                AllTXT.Add(XMLtoTXT(@"E:\Работа\май 2015\26_05_15\catalog_labirint\" + file_name, list_link));
            }

            for(int i = 0; i < AllTXT.Count; i++)
            {
                foreach (string link in AllTXT[i])
                {
                    System.IO.StreamWriter writer_lab = new System.IO.StreamWriter("txt_file_" + i.ToString() + ".txt", true, System.Text.Encoding.GetEncoding(1251));
                    writer_lab.WriteLine(link);
                    writer_lab.Close();
                }
            }
        }


        // перевод xml в txt
        public List<string> XMLtoTXT(string file, List<string> list_link)
        {
            int it = 0;
            using (XmlReader xml = XmlReader.Create(file))
            {
                while (xml.Read())
                {
                    switch (xml.NodeType)
                    {
                        case XmlNodeType.Element:
                            // нашли элемент 
                            if (xml.Name == "loc")
                            {
                                string link = xml.ReadElementContentAsString();
                                if (link.Contains("/books/"))
                                {
                                    list_link.Add(link);
                                }
                            }
                            break;
                    }

                }
            }
            return list_link;
        }

        private void AverSpeed()
        {
            //double aver_speed = double.Parse(label11.Text) / double.Parse(label13.Text);
            //label54.RefreshData(Math.Round(aver_speed, 2).ToString());
        }

        private void timerAllBase_Tick(object sender, EventArgs e)
        {
            try
            {
                OleDbConnection con = new OleDbConnection(connectionString);
                using (OleDbCommand Command = new OleDbCommand(" SELECT count (id) from labirint as total", con))
                {
                    // открываем соединение с БД с помощью метода Open() объекта OleDbConnection
                    if (con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }
                    //con.Open();
                    OleDbDataReader DB_Reader = Command.ExecuteReader();
                    if (DB_Reader.HasRows)
                    {
                        DB_Reader.Read();
                        int id = DB_Reader.GetInt32(0);
                        //label51.RefreshData(id.ToString());
                        con.Close();       // закрываем базу
                    }
                }
                AllLinkCounter();
                AverSpeed();
            }
            catch { }
            
        }

        private void button13_Click(object sender, EventArgs e)
        {
            //Request req_kazan = new Request();
            BooksKazanParser BooksKazanParse = new BooksKazanParser();

            timer_kazan_time.Start();
            timer_kazan_grid.Start();

            button13.Enabled = false;
            button11.Enabled = false;

            pictureBox3.Visible = true;
            pictureBox4.Visible = false;
            label62.Text = "0";
            label64.Text = "0";

            t_kazan = 0;

            kazanDetailTable = new System.Data.DataTable("kazanParse");

            // Define all the columns once.
            DataColumn[] cols_kazan ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            kazanDetailTable.Columns.AddRange(cols_kazan);

            kazanDetailTable.Rows.Clear();

            if (dataGridView2.Rows.Count > 0)
            {
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();
            }

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*book_kazan*"))
            {
                file.Delete();
            }

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> task_kazan = new Task<MyList<List<string>>>(() => BooksKazanParse.goParse_bookskazan(richText, ListKazan, label62, label64, timer_kazan_time)); //
                task_kazan.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button13.Enabled = true;
            }
        }

        int t_kazan = 0;
        private void timer_kazan_time_Tick(object sender, EventArgs e)
        {
            t_kazan++;
            label60.Text = t_kazan.ToString();
        }

        private void button14_Click(object sender, EventArgs e)
        {
        }

        private void timer_kazan_grid_Tick(object sender, EventArgs e)
        {
            if (timer_kazan_time.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                dataGridView2.Refresh();
                timer_kazan_grid.Stop();
                pictureBox4.Visible = true;
                pictureBox3.Visible = false;
                button11.Enabled = true;
                button12.Enabled = true;
                button13.Enabled = false;
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView2, "kazan");
        }

        private void button11_Click(object sender, EventArgs e)
        {
            dataGridView2.DataSource = null;
            dataGridView2.Rows.Clear();
            button13.Enabled = true;
            button12.Enabled = false;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            //Request req_booka = new Request();
            BookaParser BookaParse = new BookaParser();

            timer_booka.Start();
            timer_buka_grid.Start();

            pictureBox5.Visible = true;
            pictureBox6.Visible = false;

            button17.Enabled = false;
            button15.Enabled = false;

            label69.Text = "0";
            label71.Text = "0";
            label73.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*booka*"))
            {
                file.Delete();
            }

            bookaDetailTable = new System.Data.DataTable("bookaParse");

            // Define all the columns once.
            DataColumn[] cols_booka ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            bookaDetailTable.Columns.AddRange(cols_booka);

            bookaDetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> task_booka = new Task<MyList<List<string>>>(() => BookaParse.goParse_booka(richText, ListBooka, label71, label73, timer_booka)); //
                task_booka.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button17.Enabled = true;
            }
        }

        int t_booka = 0;
        private void timer_booka_Tick(object sender, EventArgs e)
        {
            t_booka++;
            label69.RefreshData(t_booka.ToString());
        }

        private void timer_buka_grid_Tick(object sender, EventArgs e)
        {
            if (timer_booka.Enabled)
            {
                //dataGridView4.Refresh();
            }
            else
            {
                dataGridView4.Refresh();
                timer_buka_grid.Stop();
                pictureBox6.Visible = true;
                pictureBox5.Visible = false;
                button15.Enabled = true;
                button16.Enabled = true;
                button17.Enabled = false;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            dataGridView4.DataSource = null;
            dataGridView4.Rows.Clear();
            button17.Enabled = true;
            button16.Enabled = false;
        }

        private void button16_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView4, "booka");
        }

        private void button20_Click(object sender, EventArgs e)
        {
            Request req_booka = new Request();

            timer_prodalit.Start();
            timer_prodalit_grid.Start();

            pictureBox7.Visible = true;
            pictureBox8.Visible = false;

            button20.Enabled = false;
            button18.Enabled = false;

            label78.Text = "0";
            label80.Text = "0";
            label82.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*prodalit*"))
            {
                file.Delete();
            }

            prodalitDetailTable = new System.Data.DataTable("prodalitParse");

            // Define all the columns once.
            DataColumn[] cols_prodalit ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            prodalitDetailTable.Columns.AddRange(cols_prodalit);

            prodalitDetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> task_prodalit = new Task<MyList<List<string>>>(() => req_booka.goParse_prodalit(richText, ListProdalit, label80, label82, timer_prodalit)); //
                task_prodalit.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button20.Enabled = true;
            }
        }

        int t_prodalit = 0;
        private void timer_prodalit_Tick(object sender, EventArgs e)
        {
            t_prodalit++;
            label78.RefreshData(t_prodalit.ToString());
        }

        private void timer_prodalit_grid_Tick(object sender, EventArgs e)
        {
            if (timer_prodalit.Enabled)
            {
                //dataGridView5.Refresh();
            }
            else
            {
                dataGridView5.Refresh();
                timer_prodalit_grid.Stop();
                pictureBox8.Visible = true;
                pictureBox7.Visible = false;
                button18.Enabled = true;
                button19.Enabled = true;
                button20.Enabled = false;
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            dataGridView5.DataSource = null;
            dataGridView5.Rows.Clear();
            button20.Enabled = true;
            button19.Enabled = false;
        }

        private void button19_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView5, "prodalit");
        }

        private void button23_Click(object sender, EventArgs e)
        {
            //Request req_100000 = new Request();
            _100000knigParser _100000Parse = new _100000knigParser();

            timer100000.Start();
            timer100000grid.Start();

            button23.Enabled = false;
            button21.Enabled = false;

            pictureBox9.Visible = true;
            pictureBox10.Visible = false;

            label87.Text = "0";
            label89.Text = "0";
            label91.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*100000knig*"))
            {
                file.Delete();
            }

            knig_100000DetailTable = new System.Data.DataTable("100000knigParse");

            // Define all the columns once.
            DataColumn[] cols_100000 ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            knig_100000DetailTable.Columns.AddRange(cols_100000);

            knig_100000DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> task_100000 = new Task<MyList<List<string>>>(() => _100000Parse.goParse100000(richText, List100000knig, label89, label91, timer100000)); //
                task_100000.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button23.Enabled = true;
            }
        }

        int t_100000 = 0;
        private void timer100000_Tick(object sender, EventArgs e)
        {
            t_100000++;
            label87.RefreshData(t_100000.ToString());
        }

        private void timer100000grid_Tick(object sender, EventArgs e)
        {
            if (timer100000.Enabled)
            {
                //dataGridView6.Refresh();
            }
            else
            {
                dataGridView6.Refresh();
                timer100000grid.Stop();
                pictureBox10.Visible = true;
                pictureBox9.Visible = false;
                button21.Enabled = true;
                button22.Enabled = true;
                button23.Enabled = false;
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            dataGridView6.DataSource = null;
            dataGridView6.Rows.Clear();
            button23.Enabled = true;
            button22.Enabled = false;
        }

        private void button22_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView6, "100000-knig");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            //Request req_mdk_arbat = new Request();
            MdkArbatParser MdkArbatParse = new MdkArbatParser();

            timer_mdk_arbat.Start();
            timer_mdk_arbat_grid.Start();

            button7.Enabled = false;
            button5.Enabled = false;

            pictureBox11.Visible = true;
            pictureBox12.Visible = false;

            label14.Text = "0";
            label16.Text = "0";
            label18.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*mdk_arbat*"))
            {
                file.Delete();
            }
            

            mdk_arbatDetailTable = new System.Data.DataTable("mdk_arbat_Parse");

            // Define all the columns once.
            DataColumn[] cols_mdk_arbat ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            mdk_arbatDetailTable.Columns.AddRange(cols_mdk_arbat);

            mdk_arbatDetailTable.Rows.Clear();
            
            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> mdk_arbat = new Task<MyList<List<string>>>(() => MdkArbatParse.goParse_mdk_arbat(richText, List_mdk_arbat, label16, label18, timer_mdk_arbat)); //
                mdk_arbat.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button7.Enabled = true;
            }
        }

        int t_mdk_arbat = 0;
        //private void timer_mdk_arbat_Tick(object sender, EventArgs e)
        //{
        //    t_mdk_arbat++;
        //    label14.RefreshData(t_mdk_arbat.ToString());
        //}

        /*
        private void timer_mdk_arbat_grid_Tick(object sender, EventArgs e)
        {
            if (timer_mdk_arbat.Enabled)
            {
                //dataGridView6.Refresh();
            }
            else
            {
                dataGridView3.Refresh();
                timer_mdk_arbat_grid.Stop();
                pictureBox12.Visible = true;
                pictureBox11.Visible = false;
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = false;
            }
        }
        */

        private void button5_Click(object sender, EventArgs e)
        {
            dataGridView3.DataSource = null;
            dataGridView3.Rows.Clear();
            button7.Enabled = true;
            button6.Enabled = false;
        }

        private void timer_mdk_arbat_Tick_1(object sender, EventArgs e)
        {
            t_mdk_arbat++;
            label14.RefreshData(t_mdk_arbat.ToString());
        }

        private void timer_mdk_arbat_grid_Tick_1(object sender, EventArgs e)
        {
            if (timer_mdk_arbat.Enabled)
            {
                //dataGridView6.Refresh();
            }
            else
            {
                dataGridView3.Refresh();
                timer_mdk_arbat_grid.Stop();
                pictureBox12.Visible = true;
                pictureBox11.Visible = false;
                button5.Enabled = true;
                button6.Enabled = true;
                button7.Enabled = false;
            }
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView3, "mdk_arbat");
        }

        private void button8_Click_1(object sender, EventArgs e)
        {
            HomeToBasket();
        }

        private void HomeToBasket()
        {
            try
            {
                using (IWebDriver driver = new ChromeDriver())
                {
                    driver.Navigate().GoToUrl("http://www.chitai-gorod.ru");

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

        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            //Request req_biblio = new Request();
            BiblioGlobusParser BiblioGlobusParse = new BiblioGlobusParser();

            timer_biblio_globus.Start();
            timer_biblio_globus_grid.Start();

            button23.Enabled = false;
            button21.Enabled = false;

            pictureBox9.Visible = true;
            pictureBox10.Visible = false;

            label87.Text = "0";
            label89.Text = "0";
            label91.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*biblo_globus*"))
            {
                file.Delete();
            }

            biblio_globus_DetailTable = new System.Data.DataTable("biblio_globusParse");

            // Define all the columns once.
            DataColumn[] cols_biblio_globus ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            biblio_globus_DetailTable.Columns.AddRange(cols_biblio_globus);

            biblio_globus_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> biblio_globus = new Task<MyList<List<string>>>(() => BiblioGlobusParse.goParse_biblio_globus(richText, List_biblio, label25, label27, timer100000)); //
                biblio_globus.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button23.Enabled = true;
            }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView7, "biblio_globus");
        }

        private void timer_biblio_globus_grid_Tick(object sender, EventArgs e)
        {
            if (timer_biblio_globus.Enabled)
            {
                //dataGridView6.Refresh();
            }
            else
            {
                dataGridView7.Refresh();
                timer_biblio_globus_grid.Stop();
                pictureBox10.Visible = true;
                pictureBox9.Visible = false;
                button9.Enabled = true;
                button10.Enabled = true;
                button14.Enabled = false;
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            //Request req_bookbars = new Request();
            BookbarsParser BookbarsParse = new BookbarsParser();


            timer_bookbars.Start();
            timer_bookbars_grid.Start();

            button23.Enabled = false;
            button21.Enabled = false;

            pictureBox9.Visible = true;
            pictureBox10.Visible = false;

            label87.Text = "0";
            label89.Text = "0";
            label91.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*bookbars*"))
            {
                file.Delete();
            }

            bookbars_DetailTable = new System.Data.DataTable("bookbarsParse");

            // Define all the columns once.
            DataColumn[] cols_bookbars ={
                                  new DataColumn("isbn",typeof(String)),
                                  new DataColumn("Наименование",typeof(String)),
                                  new DataColumn("Цена",typeof(String))
                              };

            bookbars_DetailTable.Columns.AddRange(cols_bookbars);

            bookbars_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> bookbars = new Task<MyList<List<string>>>(() => BookbarsParse.goParse_bookbars(richText, List_bookbars, label34, label36, timer_bookbars)); //
                bookbars.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button23.Enabled = true;
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            //Request req_ozon = new Request();
            OzonParser OzonParse = new OzonParser();
            timer_ozon.Start();
            timer_ozon_grid.Start();

            button29.Enabled = false;
            button27.Enabled = false;

            pictureBox17.Visible = true;
            pictureBox18.Visible = false;

            t_ozon = 0;
            label41.Text = "0";
            label43.Text = "0";
            label45.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*ozon*"))
            {
                file.Delete();
            }

            ozon_DetailTable = new System.Data.DataTable("ozonParse");

            ozon_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> ozon = new Task<MyList<List<string>>>(() => OzonParse.goParse_ozon(richText, List_ozon, label43, label45, timer_ozon, checkBox5)); //
                ozon.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button29.Enabled = true;
            }
        }

        int t_ozon = 0;
        private void timer_ozon_Tick(object sender, EventArgs e)
        {
            t_ozon++;
            label41.Text = t_ozon.ToString();
        }

        private void timer_ozon_grid_Tick(object sender, EventArgs e)
        {
            if (timer_ozon.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                //dataGridView2.Refresh();
                timer_ozon_grid.Stop();
                pictureBox18.Visible = true;
                pictureBox17.Visible = false;
                button28.Enabled = true;
                button27.Enabled = true;
                button29.Enabled = false;
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            dataGridView9.DataSource = null;
            dataGridView9.Rows.Clear();
            button29.Enabled = true;
            button27.Enabled = false;
            button28.Enabled = false;
        }

        private void button32_Click(object sender, EventArgs e)
        {
            //Request req_labirint = new Request();
            LabirintParser LabirintParse = new LabirintParser();

            timer_labirint.Start();
            timer_labirint_grid.Start();

            button32.Enabled = false;
            button30.Enabled = false;

            pictureBox19.Visible = true;
            pictureBox20.Visible = false;

            t_labirint = 0;
            label50.Text = "0";
            label52.Text = "0";
            label54.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*labirint*"))
            {
                file.Delete();
            }

            labirint_DetailTable = new System.Data.DataTable("labirintParse");

            labirint_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> labirint = new Task<MyList<List<string>>>(() => LabirintParse.goParse_labirint(richText, List_labirint, label52, label54, 
                                                                                                                            timer_labirint)); //
                labirint.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button32.Enabled = true;
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            //Request req_read = new Request();
            ReadParser ReadParse = new ReadParser();

            timer_read.Start();
            timer_read_grid.Start();

            button34.Enabled = false;
            button35.Enabled = false;

            pictureBox21.Visible = true;
            pictureBox22.Visible = false;

            t_read = 0;
            label95.Text = "0";
            label97.Text = "0";
            label99.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*read*"))
            {
                file.Delete();
            }

            read_DetailTable = new System.Data.DataTable("readParse");

            read_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> read = new Task<MyList<List<string>>>(() => ReadParse.goParse_read(richText, List_read, label97, label99, timer_read)); //
                read.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button32.Enabled = true;
            }
        }

        int t_labirint = 0;
        private void timer_labirint_Tick(object sender, EventArgs e)
        {
            t_labirint++;
            label50.Text = t_labirint.ToString();
        }

        private void timer_labirint_grid_Tick(object sender, EventArgs e)
        {
            if (timer_labirint.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                //dataGridView2.Refresh();
                timer_labirint_grid.Stop();
                pictureBox20.Visible = true;
                pictureBox19.Visible = false;
                button30.Enabled = true;
                button31.Enabled = true;
                button32.Enabled = false;
            }
        }

        int t_read = 0;
        private void timer_read_Tick(object sender, EventArgs e)
        {
            t_read++;
            label95.Text = t_read.ToString();
        }

        private void timer_read_grid_Tick(object sender, EventArgs e)
        {
            if (timer_read.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                //dataGridView2.Refresh();
                timer_read_grid.Stop();
                pictureBox22.Visible = true;
                pictureBox21.Visible = false;
                button33.Enabled = true;
                button34.Enabled = true;
                button35.Enabled = false;
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            dataGridView11.DataSource = null;
            dataGridView11.Rows.Clear();
            button35.Enabled = true;
            button34.Enabled = false;
        }

        private void button30_Click(object sender, EventArgs e)
        {
            dataGridView10.DataSource = null;
            dataGridView10.Rows.Clear();
            button32.Enabled = true;
            button31.Enabled = false;
        }

        private void button28_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView9, "ozon");
        }

        private void button31_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView10, "labirint");
        }

        private void button34_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView11, "read");
        }

        private void button38_Click(object sender, EventArgs e)
        {
            //Request req_kniga = new Request();
            KnigaParse KnigaParse = new KnigaParse();

            timer_kniga.Start();
            timer_kniga_grid.Start();

            button36.Enabled = false;
            button37.Enabled = false;

            pictureBox23.Visible = true;
            pictureBox24.Visible = false;

            t_kniga = 0;
            label108.Text = "0";
            label110.Text = "0";
            label112.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*kniga*"))
            {
                file.Delete();
            }

            kniga_DetailTable = new System.Data.DataTable("knigaParse");

            kniga_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> kniga = new Task<MyList<List<string>>>(() => KnigaParse.goParse_kniga(richText, List_kniga, label110, label112, timer_kniga)); //
                kniga.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button32.Enabled = true;
            }
        }

        int t_kniga = 0;
        private void timer_kniga_Tick(object sender, EventArgs e)
        {
            t_kniga++;
            label108.Text = t_kniga.ToString();
        }

        private void timer_kniga_grid_Tick(object sender, EventArgs e)
        {
            if (timer_kniga.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                //dataGridView2.Refresh();
                timer_kniga_grid.Stop();
                pictureBox24.Visible = true;
                pictureBox23.Visible = false;
                button36.Enabled = true;
                button37.Enabled = true;
                button38.Enabled = false;
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            dataGridView12.DataSource = null;
            dataGridView12.Rows.Clear();
            button38.Enabled = true;
            button37.Enabled = false;
        }

        private void button37_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView12, "kniga");
        }

        private void button41_Click(object sender, EventArgs e)
        {
            //Request req_biblion = new Request();
            BiblionParser BibParser = new BiblionParser();
            timer_biblion.Start();
            timer_biblion_grid.Start();

            button39.Enabled = false;
            button40.Enabled = false;

            pictureBox25.Visible = true;
            pictureBox26.Visible = false;

            t_biblion = 0;
            label108.Text = "0";
            label110.Text = "0";
            label112.Text = "0";

            DirectoryInfo dir = new DirectoryInfo(Directory.GetCurrentDirectory());
            foreach (FileInfo file in dir.GetFiles("*biblion*"))
            {
                file.Delete();
            }

            biblion_DetailTable = new System.Data.DataTable("biblionParse");

            biblion_DetailTable.Rows.Clear();

            List<string> richText = new List<string>();
            if (richTextBox1.Lines.Count() > 1)
            {
                foreach (string str_rich in richTextBox1.Lines)
                {
                    richText.Add(str_rich);
                }

                Task<MyList<List<string>>> biblion = new Task<MyList<List<string>>>(() => BiblionParser.goParse_biblion(richText, List_biblion, label120, label122, timer_biblion)); //
                biblion.Start();
            }
            else
            {
                MessageBox.Show("Вы не выбрали файл *.csv для поиска значений!");
                button32.Enabled = true;
            }
        }

        int t_biblion = 0;
        private void timer_biblion_Tick(object sender, EventArgs e)
        {
            t_biblion++;
            label118.Text = t_biblion.ToString();
        }

        private void timer_biblion_grid_Tick(object sender, EventArgs e)
        {
            if (timer_biblion.Enabled)
            {
                //dataGridView2.Refresh();
            }
            else
            {
                //dataGridView2.Refresh();
                timer_biblion_grid.Stop();
                pictureBox26.Visible = true;
                pictureBox25.Visible = false;
                button39.Enabled = true;
                button40.Enabled = true;
                button41.Enabled = false;
            }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            dataGridView13.DataSource = null;
            dataGridView13.Rows.Clear();
            button41.Enabled = true;
            button40.Enabled = false;
        }

        private void button40_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView13, "biblion");
        }

        int t_bookbars = 0;
        private void timer_bookbars_Tick(object sender, EventArgs e)
        {
            t_bookbars++;
            label32.Text = t_bookbars.ToString();
        }

        private void timer_bookbars_grid_Tick(object sender, EventArgs e)
        {
            if (!timer_bookbars.Enabled)
            {
                timer_bookbars_grid.Stop();
                pictureBox16.Visible = true;
                pictureBox15.Visible = false;
                button24.Enabled = true;
                button25.Enabled = true;
                button26.Enabled = false;
            }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            ExportToExcel(dataGridView8, "bookbars");
        }

        private void button24_Click(object sender, EventArgs e)
        {
            dataGridView8.DataSource = null;
            dataGridView8.Rows.Clear();
            button26.Enabled = true;
            button24.Enabled = false;
            button25.Enabled = false;
        }

        int t_biblioglob = 0;
        private void timer_biblio_globus_Tick(object sender, EventArgs e)
        {
            t_biblioglob++;
            label23.Text = t_biblioglob.ToString();
        }
    }


    public class MyList<T> : List<T>
    {
        public event EventHandler OnAdd;

        public void Add(T item)
        {
            base.Add(item);
            if (null != OnAdd)
            {
                OnAdd(this, null);
            }
        }
    }

    static class Ext
    {
        delegate void SetTable(System.Data.DataTable table);
        public static void SetTableIn(this DataGridView dgv, System.Data.DataTable tbl)
        {
            dgv.Invoke(new SetTable(t => { dgv.DataSource = t; }), tbl);
        }

        delegate void RefreshLabel(string text);
        public static void RefreshData(this System.Windows.Forms.Label label, string text)
        {
            label.Invoke(new RefreshLabel(t => { label.Text = t; }), text);
        }
    }
}
