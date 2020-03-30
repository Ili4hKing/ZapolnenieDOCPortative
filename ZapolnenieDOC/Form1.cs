using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using DataTable = System.Data.DataTable;
using System.Data.Entity;

namespace ZapolnenieDOC
{
    public partial class Form1 : Form
    {
        private Application application;
        private Workbook workBook;
        private Worksheet worksheet;




        public Form1()
        {
            InitializeComponent();
            using (IDbConnection connection = new SqlConnection())
            {

            }
        }
        public struct Person
        {
            public string FIO
            { get; set; }
            public DateTime DateBirdhsday
            { get; set; }

        }

        public struct NekorektData
        {
            public string FIO
            { get; set; }
            public DateTime DateBirdhsday
            { get; set; }

        }

        public struct ДанныеПоТаблицеШаблоныГруппы
        {
            public string ФИО { get; set; }
            public System.DateTime ДатаРождения { get; set; }
            public string МестоРождения { get; set; }
            public string АдресПоРегистрации { get; set; }
            public string Телефон { get; set; }
            public string Паспорт { get; set; }
            public string Email { get; set; }
            public int id { get; set; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            var someList = new List<string>();
            DataTable dataTable = new DataTable();
            List<Person> persons = new List<Person>();

            foreach (DataGridViewRow row2 in dataGridView1.Rows)
            {

                string searchValue = row2.Cells[1].Value.ToString();
                string tr = row2.Cells[6].Value.ToString();

                string tlFio = searchValue;
                string[] c = tlFio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string q = c[0] + " " + c[1] + " " + c[2];

                DateTime d = Convert.ToDateTime(row2.Cells[2].Value.ToString());


                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                try
                {
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        string rr = row.Cells[1].Value.ToString();
                        string tlFio2 = rr;
                        string[] r = tlFio2.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        string p = r[0] + " " + r[1] + " " + r[2];

                        DateTime t = Convert.ToDateTime(row.Cells[2].Value.ToString());

                        if (p.Equals(q) && t == d)
                        {


                            if (string.IsNullOrEmpty(tr) || tr == " ")
                            {
                                row2.Cells[6].Value = row.Cells[3].Value.ToString();
                                Person person = new Person
                                {
                                    FIO = searchValue,
                                    DateBirdhsday = d

                                };
                                persons.Add(person);
                            }

                        }
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }




            }
            foreach (DataGridViewRow row22 in dataGridView1.Rows)
            {

                string searchValue = row22.Cells[1].Value.ToString();
                string tr = row22.Cells[6].Value.ToString();
                DateTime d = Convert.ToDateTime(row22.Cells[2].Value.ToString());

                List<NekorektData> nekorektDatas = new List<NekorektData>();

                if (string.IsNullOrEmpty(tr) || tr == " ")
                {
                    NekorektData nekorektData = new NekorektData
                    {
                        FIO = searchValue,
                        DateBirdhsday = d
                    };
                    nekorektDatas.Add(nekorektData);
                }
                if (nekorektDatas.Count > 0)
                {
                    foreach (NekorektData l in nekorektDatas)
                    {
                        listBox2.Items.Add("ФИО " + l.FIO + " Дата рождения " + l.DateBirdhsday);
                    }

                }

                for (int i = 0; i < dataGridView1.Columns.Count && someList.Count != dataGridView1.Columns.Count; i++)
                {
                    dataTable.Columns.Add();
                    string it = dataGridView1.Columns[i].HeaderText;
                    //someList.Add(it.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", ""));
                    someList.Add(it);

                }
                string[] rrt = new string[someList.Count];
                for (int r = 0; r < someList.Count; r++)
                {

                    rrt[r] = row22.Cells[r].Value.ToString();


                }
                dataTable.Rows.Add(rrt);


            }

            if (persons.Count > 0)
            {
                foreach (Person p in persons)
                {
                    listBox1.Items.Add("ФИО " + p.FIO + " Дата рождения " + p.DateBirdhsday);
                }

            }


            dataGridView3.DataSource = dataTable;

            for (int i = 0; i < someList.Count; i++)
                dataGridView3.Columns[i].HeaderText = someList[i];


        }

        private void button2_Click(object sender, EventArgs e)
        {

            //string yourtext = "18.02. 2003";
            //string tlFio = yourtext;
            //string[] b = tlFio.Split(new char[] { ' ', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
            //string o = b[0] + "." + b[1] + "." + b[2];
            ////string text = yourtext.Replace(" ", ".");
            //DateTime d = Convert.ToDateTime(o);
            ////DateTime.TryParseExact(yourtext, "0:MM/dd/yy H:mm:ss zzz", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out d);

            //this.Hide();
            //Form2 rr = new Form2();
            //rr.Show();





            //foreach (DataGridViewRow Row in dataGridView1.Rows)

            //{

            //    listBox1.Items.Add(Row.AccessibilityObject.Value);

            //    //foreach (DataGridViewCell Cell in dataGridView1.Cell)
            //    //{


            //    //}

            //}
            ////dataGridView1.Rows[1].Cells[1].Value
            ///

            var someList = new List<string>();
            DataTable dataTable = new DataTable();
            List<Person> persons = new List<Person>(); 

            if (dataGridView1.Rows.Count > 0 && dataGridView1.Columns.Count > 0) // Разобраться почему не проходит условие по дате и фио так как вставляеться данные рандомные 
            {



                for (int i = 0; i < dataGridView1.Columns.Count && someList.Count != dataGridView1.Columns.Count; i++)
                {
                    dataTable.Columns.Add();
                    string it = dataGridView1.Columns[i].HeaderText;
                    //someList.Add(it.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", ""));
                    someList.Add(it);

                }
                for (int i = 0; i < dataGridView1.Rows.Count ; i++)
                {
                    string[] row = new string[dataGridView1.Columns.Count];
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        string tt = Convert.ToString(dataGridView1[j , i ].Value);
                        row[j] = tt.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", "");
                    }
                    string tlFio = row[1];
                    string[] b = tlFio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string o = b[0] + " " + b[1] + " " + b[2];

                    if (dataGridView2.Rows.Count > 0 && dataGridView2.Columns.Count > 0)
                    {

                        for (int r = 0; r < dataGridView2.Rows.Count; r++)
                        {
                            string[] row2 = new string[dataGridView2.Columns.Count];
                            for (int h = 0; h < dataGridView2.Columns.Count; h++)
                            {
                                string tt = Convert.ToString(dataGridView2[h, r].Value);
                                row2[h] = tt.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", "");
                            }
                            string rlFio = row2[1];
                            string[] c = tlFio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            string q = c[0] + " " + c[1] + " " + c[2];

                            if (o == q && row[2]==row2[2])
                            {
                                if (string.IsNullOrEmpty(row[6]) || row[6] == " ")
                                {
                                    row[6] = row2[3];
                                    Person person = new Person
                                    {
                                        FIO = row[1],
                                        DateBirdhsday = Convert.ToDateTime(row[2])

                                    };
                                    persons.Add(person);
                                }



                            }

                            
                        }


                    }


                    dataTable.Rows.Add(row);
                }


            }

            if (persons.Count > 0)
            {
                foreach (Person p in persons)
                {
                    listBox1.Items.Add("ФИО " + p.FIO + " Дата рождения " + p.DateBirdhsday);
                }

            }

            

                dataGridView3.DataSource = dataTable;

                for (int i = 0; i < someList.Count; i++)
                    dataGridView3.Columns[i].HeaderText = someList[i];




            if (dataGridView3.Rows.Count > 0 && dataGridView3.Columns.Count > 0)
            {

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    string[] row = new string[dataGridView3.Columns.Count];
                    for (int j = 0; j < dataGridView3.Columns.Count; j++)
                    {
                        string tt = Convert.ToString(dataGridView3[j, i].Value);
                        row[j] = tt.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", "");
                    }



                    List<NekorektData> nekorektDatas = new List<NekorektData>();

                    if (string.IsNullOrEmpty(row[6]) || row[6] == " ")
                    {
                        NekorektData nekorektData = new NekorektData
                        {
                            FIO = row[1],
                            DateBirdhsday = Convert.ToDateTime(row[2])
                        };
                        nekorektDatas.Add(nekorektData);
                    }
                    if (nekorektDatas.Count > 0)
                    {
                        foreach (NekorektData l in nekorektDatas)
                        {
                            listBox2.Items.Add("ФИО " + l.FIO + " Дата рождения " + l.DateBirdhsday);
                        }

                    }
                }
            }


            }

        private void button3_Click(object sender, EventArgs e)
        {
            object missing = Type.Missing;

            
                
                int t = 0;
                int i = 0;
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    t++;



                }

                // Открываем приложение
                application = new Application



                {
                    DisplayAlerts = false
                };


                // Файл шаблона
                //const string template = "E:\\Shablone.xlsx";

                // Открываем книгу
                workBook = application.Workbooks.Add();

                // Получаем активную таблицу
                worksheet = workBook.ActiveSheet as Worksheet;

                // Записываем данные
                worksheet.Range["A1"].Value = "№";
                worksheet.Range["B1"].Value = "ФИО";
                worksheet.Range["C1"].Value = "Дата рождения";
                worksheet.Range["D1"].Value = "Место рождения";
                worksheet.Range["E1"].Value = "Адрес по регистрации";
                worksheet.Range["F1"].Value = "Телефон";
                worksheet.Range["G1"].Value = "Паспорт";
                worksheet.Range["H1"].Value = "Email";
                


                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    //t++;

                    if (i < t)
                    {   DateTime date = Convert.ToDateTime(row.Cells[2].Value.ToString());
                        string s1 = date.ToString("MM/dd/yyyy");
                        i++;
                        worksheet.Cells[i + 1, 1].Value = row.Cells[0].Value.ToString(); 
                        worksheet.Cells[i + 1, 2].Value = row.Cells[1].Value.ToString();
                        worksheet.Cells[i + 1, 3].Value = s1;
                        worksheet.Cells[i + 1, 4].Value = row.Cells[3].Value.ToString(); 
                        worksheet.Cells[i + 1, 5].Value = row.Cells[4].Value.ToString();
                        worksheet.Cells[i + 1, 6].Value = row.Cells[5].Value.ToString();
                        worksheet.Cells[i + 1, 7].Value = row.Cells[6].Value.ToString();
                        worksheet.Cells[i + 1, 8].Value = row.Cells[7].Value.ToString();

                    }
                }
                // Показываем приложение
                application.Visible = true;
                TopMost = true;
                object template3 = "E:\\Shablone" + ".xlsx";
                string savedFileName = textBox3.Text+"\\ШаблоныГруппВыгрузкаССервера.xlsx"; //Добавить возможность выбора куда сохранять
                workBook.SaveAs(Path.Combine(Environment.CurrentDirectory, savedFileName));

                CloseExcel();
                MessageBox.Show("Файл сохранен путь: "+savedFileName);
        }
        
        private void CloseExcel()
        {
            if (application != null)
            {
                int excelProcessId = -1;
                GetWindowThreadProcessId(application.Hwnd, ref excelProcessId);

                Marshal.ReleaseComObject(worksheet);
                workBook.Close();
                Marshal.ReleaseComObject(workBook);
                application.Quit();
                Marshal.ReleaseComObject(application);

                application = null;
                // Прибиваем висящий процесс
                try
                {
                    Process process = Process.GetProcessById(excelProcessId);
                    process.Kill();
                }
                finally { }
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(int hWnd, ref int lpdwProcessId);

        private void button4_Click(object sender, EventArgs e)
        {
            using (TexnikymBDEntities db = new TexnikymBDEntities())
            {
                object missing = Type.Missing;


                Object Pa = textBox1.Text; // Путь к шаблону 

                Word.Application wordApp = new Word.Application();// Создаём объект приложения


                wordApp.Documents.Open(ref Pa, ref missing, true, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);// доделать чтобы не вылетало сообщение

                Word.Document document = wordApp.ActiveDocument;

                int countTable = document.Tables.Count;

                var someList = new List<string>();
                DataTable dataTable = new DataTable();

                for (int y = 1; y <= countTable; y++)
                {



                    List<ШаблонГруппы> ShabloniGr = new List<ШаблонГруппы>();
                    Word.Table table = document.Tables[y];
                    

                    if (table.Rows.Count > 0 && table.Columns.Count > 0)
                    {
                       

                        for (int i = 0; i < table.Columns.Count && someList.Count != table.Columns.Count; i++)
                        {
                            dataTable.Columns.Add();
                            someList.Add(table.Cell(1, i + 1).Range.Text.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", ""));

                        }

                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            progressBar1.Maximum = (countTable ) * table.Columns.Count;
                            progressBar1.Value++;


                        }
                        for (int i = 0; i < table.Rows.Count - 1; i++)
                        {
                            string[] row = new string[table.Columns.Count];
                            for (int j = 0; j < table.Columns.Count; j++)
                                row[j] = table.Cell(i + 2, j + 1).Range.Text.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", "");



                            DateTime d;
                            string dateConv = row[2].Replace("\a", "");
                            string tlFio = dateConv;
                            string[] b = tlFio.Split(new char[] { ' ', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            string o = b[0] + "." + b[1] + "." + b[2];

                            if (DateTime.TryParse(o, out d))

                                d = Convert.ToDateTime(o);
                            else
                                o = "2000-01-01 00:00:00.000";// Если дата введена не коретно то вводиться это число
                            d = Convert.ToDateTime(o);

                            row[2] = Convert.ToString(d);

                            dataTable.Rows.Add(row);
                        }
                       

                    }
                   

                }
                dataGridView1.DataSource = dataTable;
                for (int i = 0; i < someList.Count; i++)
                    dataGridView1.Columns[i].HeaderText = someList[i];


                wordApp.ActiveDocument.Close();
                wordApp.Quit();

                //db.ШаблонГруппы.Load();

                //dataGridView1.DataSource = db.ШаблонГруппы.Local.ToBindingList();

                MessageBox.Show("Данные помещены");



            }

        }



        private void textBox1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "MS Word 2007 (*.docx)|*.docx|MS Word 2003 (*.doc)|*.doc";
            dialog.Title = "Выберите документ для загрузки данных";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = dialog.FileName;
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (TexnikymBDEntities db = new TexnikymBDEntities())
            {
                object missing = Type.Missing;


                Object Pa = textBox2.Text; // Путь к шаблону 

                Word.Application wordApp = new Word.Application();// Создаём объект приложения


                wordApp.Documents.Open(ref Pa, ref missing, true, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing);

                Word.Document document = wordApp.ActiveDocument;

                int countTable = document.Tables.Count;
                var someList = new List<string>();
                DataTable dataTable = new DataTable();

                for (int y = 1; y <= countTable; y++)
                {

                    List<Студенты2> ShabloniGr = new List<Студенты2>();
                    Word.Table table = document.Tables[y];

                    if (table.Rows.Count > 0 && table.Columns.Count > 0)
                    {

                        for (int i = 0; i < table.Columns.Count && someList.Count != table.Columns.Count; i++)
                        {
                            dataTable.Columns.Add();
                            someList.Add(table.Cell(1, i + 1).Range.Text.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", ""));

                        }


                        for (int i = 0; i < table.Columns.Count; i++)
                        {
                            progressBar2.Maximum = (countTable) * table.Columns.Count;
                            progressBar2.Value++;


                        }
                        for (int i = 0; i < table.Rows.Count - 1; i++)
                        {
                            string[] row = new string[table.Columns.Count];
                            for (int j = 0; j < table.Columns.Count; j++)
                                row[j] = table.Cell(i + 2, j + 1).Range.Text.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", "");



                            DateTime d;
                            string dateConv = row[2].Replace("\a", "");
                            string tlFio = dateConv;
                            string[] b = tlFio.Split(new char[] { ' ', '.', ',' }, StringSplitOptions.RemoveEmptyEntries);
                            string o = b[0] + "." + b[1] + "." + b[2];

                            if (DateTime.TryParse(o, out d))

                                d = Convert.ToDateTime(o);
                            else
                                o = "2000-01-01 00:00:00.000";// Если дата введена не коретно то вводиться это число 2000-01-01 00:00:00.000
                            d = Convert.ToDateTime(o);

                            row[2] = Convert.ToString(d);

                            dataTable.Rows.Add(row);




                        }

                    }
                    

                    
                }

                dataGridView2.DataSource = dataTable;
                for (int i = 0; i < someList.Count; i++) {
                    //dataGridView1.Columns.Add(someList[i], "");
                    dataGridView2.Columns[i].HeaderText = someList[i];
                }



                wordApp.ActiveDocument.Close();
                wordApp.Quit();

               

                MessageBox.Show("Данные помещены");

            }
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "MS Word 2007 (*.docx)|*.docx|MS Word 2003 (*.doc)|*.doc";
            dialog.Title = "Выберите документ для загрузки данных";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox2.Text = dialog.FileName;
            }
        }

        private void button6_Click(object sender, EventArgs e)

        {
            var someList = new List<string>();
            DataTable dataTable = new DataTable();
            List<Person> persons = new List<Person>();

            foreach (DataGridViewRow row2 in dataGridView1.Rows)
            {

                string searchValue = row2.Cells[1].Value.ToString();
                string tr = row2.Cells[6].Value.ToString();

                string tlFio = searchValue;
                string[] c = tlFio.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                string q = c[0] + " " + c[1] + " " + c[2];  

                DateTime d = Convert.ToDateTime(row2.Cells[2].Value.ToString());


                dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                try
                {
                    foreach (DataGridViewRow row in dataGridView2.Rows)
                    {
                        string rr = row.Cells[1].Value.ToString();
                        string tlFio2 = rr;
                        string[] r = tlFio2.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        string p = r[0] + " " + r[1] + " " + r[2];

                        DateTime t = Convert.ToDateTime(row.Cells[2].Value.ToString());

                        if (p.Equals(q) && t == d)
                        {


                            if (string.IsNullOrEmpty(tr) || tr == " ")
                            {
                                row2.Cells[6].Value = row.Cells[3].Value.ToString();
                                Person person = new Person
                                {
                                    FIO = searchValue,
                                    DateBirdhsday = d

                                };
                                persons.Add(person);
                            }
                            
                        }
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }

                


            }
            foreach (DataGridViewRow row22 in dataGridView1.Rows)
            {

                string searchValue = row22.Cells[1].Value.ToString();
                string tr = row22.Cells[6].Value.ToString();
                DateTime d = Convert.ToDateTime(row22.Cells[2].Value.ToString());

                List<NekorektData> nekorektDatas = new List<NekorektData>();

                if (string.IsNullOrEmpty(tr) || tr == " ")
                {
                    NekorektData nekorektData = new NekorektData
                    {
                        FIO = searchValue,
                        DateBirdhsday = d
                    };
                    nekorektDatas.Add(nekorektData);
                }
                if (nekorektDatas.Count > 0)
                {
                    foreach (NekorektData l in nekorektDatas)
                    {
                        listBox2.Items.Add("ФИО " + l.FIO + " Дата рождения " + l.DateBirdhsday);
                    }

                }

                for (int i = 0; i < dataGridView1.Columns.Count && someList.Count != dataGridView1.Columns.Count; i++)
                {
                    dataTable.Columns.Add();
                    string it = dataGridView1.Columns[i].HeaderText;
                    //someList.Add(it.Trim('a', 'r', 'n', 't').Replace("\r", " ").Replace("\a", "").Replace("\t", ""));
                    someList.Add(it);

                }
                string[] rrt = new string[someList.Count];
                for (int r = 0; r < someList.Count; r++) {

                    rrt[r] = row22.Cells[r].Value.ToString();


                }
                dataTable.Rows.Add(rrt);
                

            }

            if (persons.Count > 0)
            {
                foreach (Person p in persons)
                {
                    listBox1.Items.Add("ФИО " + p.FIO + " Дата рождения " + p.DateBirdhsday);
                }

            }


            dataGridView3.DataSource = dataTable;

            for (int i = 0; i < someList.Count; i++)
                dataGridView3.Columns[i].HeaderText = someList[i];


        }

            private void textBox3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog DirDialog = new FolderBrowserDialog();
            DirDialog.Description = "Выбор директории";
            DirDialog.SelectedPath = @"C:\";

            if (DirDialog.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = DirDialog.SelectedPath;
            }
        }

        private void label6_Click(object sender, EventArgs e)
        {

        }
    }
}
