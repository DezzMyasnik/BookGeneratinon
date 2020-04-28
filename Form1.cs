using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelObj = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Threading;
namespace BookGeneratinon
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NwSheet;
            ExcelObj.Range ShtRange;
            DataTable dt = new DataTable();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
               // textBox1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName);

                //Устанавливаем номер листа из котрого будут извлекаться данные
                //Листы нумеруются от 1
                NwSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                ShtRange = NwSheet.UsedRange;

                for (int Cnum = 1; Cnum <= dataGridView1.Columns.Count; Cnum++)
                {
                    dt.Columns.Add(dataGridView1.Columns[Cnum].HeaderText);
                }
                dt.AcceptChanges();

                //string[] columnNames = new String[dt.Columns.Count];
                //for (int i = 0; i < dt.Columns.Count; i++)
                //{
                //    columnNames[0] = dt.Columns[i].ColumnName;
                //}

                for (int Rnum = 2; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum-1] = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                   dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                dataSet1.Tables.Add(dt);
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView1.DataSource = dataSet1.Tables[0];

                //dataGridView1.DataMember = 
                dataGridView1.Refresh();
                workbook.Close();
                app.Quit();
            }
            else
                Application.Exit();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            progressBar1.Maximum = dataSet1.Tables[0].Rows.Count;
            backgroundWorker1.RunWorkerAsync();
            //doc_res.Close();
            //app_res.Quit();

        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string result = string.Empty;
            Word.Application app_res = new Word.Application();
            Word.Document doc_res = new Word.Document();
            app_res.Visible = false;
            
            app_res.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            object nullobject = System.Reflection.Missing.Value;
            Object pathToSaveObj = Directory.GetCurrentDirectory() + "\\Result.docx";
            doc_res.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocumentDefault, ref nullobject, ref nullobject, ref nullobject,
                ref nullobject, false, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,
                ref nullobject, ref nullobject, ref nullobject, ref nullobject);
            
            if (!File.Exists(Directory.GetCurrentDirectory() + "\\Result.docx"))
            {
                doc_res = app_res.Documents.Add(ref nullobject, ref nullobject, ref nullobject, ref nullobject);
            }
            else
            {
                doc_res = app_res.Documents.Open(Directory.GetCurrentDirectory() + "\\Result.docx");
            }
            vOutput("Создали выходной файл " + Directory.GetCurrentDirectory() + "\\Result.docx"+Environment.NewLine);
            int i = 0;
            foreach (DataRow row in dataSet1.Tables[0].Rows)
            {
          
                string filename = string.Format("orig{0}", row.ItemArray[0].ToString());

                string[] f = Directory.GetFiles(Directory.GetCurrentDirectory() + "\\docs", filename + "*");

                //Word.Application wordObject = new Word.Application();
                object File = f[0];

                Word.Application wordobject = new Word.Application();
                wordobject.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                Word.Document docs= new Word.Document();
                //wordobject.Visible = false;
                try
                {
                    // docs = wordobject.Documents.Open(ref File, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject);

                    //docs.ActiveWindow.Selection.WholeStory();
                    //docs.ActiveWindow.Selection.Copy();
                    //Thread.Sleep(500);
                    // object temp = Clipboard.GetData("");
                    Microsoft.Office.Interop.Word.Paragraph para1 = doc_res.Content.Paragraphs.Add(ref nullobject);
                    // object styleHeading1 = "Заголовок 2";
                    //para1.Range.set_Style(styleHeading1);
                    para1.Range.InsertParagraph();
                    para1.Range.Font.Size = 14;
                    para1.Range.Font.Name = "Times New Roman";
                    para1.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    para1.Range.Text = row.ItemArray[1].ToString();
                    para1.Range.InsertParagraphAfter();
                    doc_res.Save();
                    para1.Range.Font.Size = 12;
                    para1.Range.Font.Name = "Times New Roman";
                    para1.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    //para1.Range.Text = row.ItemArray[1].ToString();
                    para1.Range.Text = row.ItemArray[3].ToString();
                    para1.Range.InsertParagraphAfter();
                    doc_res.Save();


                    para1.Range.Font.Size = 16;
                    para1.Range.Font.Name = "Times New Roman";
                    para1.Format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    para1.Range.Font.Bold = 1;
                    para1.Range.Text = row.ItemArray[2].ToString();

                    para1.Range.InsertParagraphAfter();
                    doc_res.Save();

                    ////object missing = Missing.Value;
                    //object what = Word.WdGoToItem.wdGoToLine;
                    //object which = Word.WdGoToDirection.wdGoToLast;
                    //Word.Range endRange = doc_res.GoTo(ref what, ref which, ref nullobject, ref nullobject);
                    //para1.Range.Paste();
                    app_res.Selection.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast, ref nullobject, ref nullobject);
                    app_res.Selection.InsertFile(f[0]);
                    //doc_res.ActiveWindow.Selection.Paste();
                    Thread.Sleep(200);

                    doc_res.Save();
                    string output = string.Format("Добавили файл {0} в выходной файл", f[0]);
                    vOutput(output + Environment.NewLine);
                    //richTextBox1.Text += Environment.NewLine + Environment.NewLine + row.Cells[1].Value.ToString() + Environment.NewLine + row.Cells[3].Value.ToString() + Environment.NewLine + row.Cells[2].Value.ToString() + Environment.NewLine;
                    // richTextBox1.Paste();
                    // rtfMain.Document.Paste();
                   // Thread.Sleep(200);
                    docs.Close(ref nullobject, ref nullobject, ref nullobject);
                    //Thread.Sleep(200);
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                   // Thread.Sleep(200);

                    backgroundWorker1.ReportProgress(i);
                    i++;
                }
                catch (Exception error)
                {
                    docs.Close();
                    wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                   // Thread.Sleep(200);
                    //wordobject.Quit(ref nullobject, ref nullobject, ref nullobject);
                    vOutput("Проблема с файлом: " + f[0] + "   " + error.Message+ Environment.NewLine);
                    backgroundWorker1.ReportProgress(i);
                    i++;
                    //return;
                    //throw error;
                }

            }
            //Object pathToSaveObj = Directory.GetCurrentDirectory() + "\\Result.docx";
            //doc_res.SaveAs(ref pathToSaveObj, Word.WdSaveFormat.wdFormatDocumentDefault, ref nullobject, ref nullobject, ref nullobject, 
            //    ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject, ref nullobject,
            //    ref nullobject, ref nullobject, ref nullobject, ref nullobject);
            app_res.Visible = true;
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            dataGridView1.Rows[e.ProgressPercentage].Selected = true;
        }
        delegate void SetTextCallback(string text);
        public void vOutput(string signaltime)
        {
            try
            {
                if (this.richTextBox1.InvokeRequired)
                {
                    SetTextCallback d = new SetTextCallback(vOutput);
                    this.Invoke(d, new object[] { signaltime });
                }
                else
                {
                    this.richTextBox1.AppendText(signaltime);
                    



                    this.richTextBox1.SelectionStart = richTextBox1.Text.Length;
                    this.richTextBox1.ScrollToCaret();
                }
            }
            catch (SystemException ex)
            {
                string es = ex.Message;
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            progressBar1.Visible = false;
        }
    }
}
