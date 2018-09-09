using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using IronPdf;
using System.Threading;

namespace FastPDFScraper
{
    public partial class FastPDFScraper : Form
    {
        private string[] pdfFiles = null;
        private string keyFile = null;
        List<Record> result = new List<Record>();
        private int totalPages = 0;
        private int currentPage = 0;
        public FastPDFScraper()
        {
            InitializeComponent();
        }

        private void btnOpenPDFFolder_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == DialogResult.OK)
            {
                pdfFiles = Directory.GetFiles(fbd.SelectedPath);
            }
            if (pdfFiles == null) return;
            totalPages = 0;
            for(int i = 0;i < pdfFiles.Length; i++)
            {
                PdfDocument PDF = PdfDocument.FromFile(pdfFiles[i]);
                totalPages += PDF.PageCount;
            }
            currentPage = 0;
        }

        private void btnOpenKeywordFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV Files|*.csv";
            DialogResult result = ofd.ShowDialog();

            if (result == DialogResult.OK)
            {
                keyFile = ofd.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                if (pdfFiles == null || pdfFiles.Length == 0 || keyFile == null) return;
                string text = System.IO.File.ReadAllText(keyFile);
                string[] keys = text.Split(',');
                for (int i = 0; i < pdfFiles.Length; i++)
                {
                    Console.WriteLine("i=" + i);
                    PdfDocument pdf = PdfDocument.FromFile(pdfFiles[i]);
                    for (int j = 0; j < pdf.PageCount; j++)
                    {
                        Console.WriteLine("j=" + j);
                        Parameter param = new Parameter()
                        {
                            Keys = keys,
                            FileName = pdfFiles[i],
                            Page = j,
                            PDF = pdf
                        };
                        Thread thread = new Thread(new ParameterizedThreadStart(search));
                        thread.Start(param);
                    }
                }
                backgroundWorker.WorkerReportsProgress = true;
                backgroundWorker.RunWorkerAsync();
            }
            catch(Exception ex)
            {

            }
        }

        void search(object param)
        {
            try
            {
                Parameter p = (Parameter)param;
                Console.WriteLine("page=" + p.Page);
                string Text = p.PDF.ExtractTextFromPage(p.Page);
                if (Text == null) return;
                for (int i = 0; i < p.Keys.Length; i++)
                    if (Text.Contains(p.Keys[i]))
                        result.Add(new Record() { Key = p.Keys[i], FileName = p.FileName, Page = p.Page });
                currentPage++;
            } catch (Exception ex)
            {
            }
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int percents = (currentPage * 100) / totalPages;
            backgroundWorker.ReportProgress(percents);
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Finish");
            using (StreamWriter writetext = new StreamWriter("result.csv"))
            {
                writetext.WriteLine("Key,FileName,PageNo");
                for ( int i = 0; i < result.Count; i++)
                {
                    writetext.WriteLine(result[i].Key + "," + result[i].FileName + "," + result[i].Page);
                }
            }
        }
    }
    public class Parameter
    {
        public string[] Keys { get; set; }
        public string FileName { get; set; }
        public int Page { get; set; }
        public PdfDocument PDF { get; set; }
    }

    public class Record : IEquatable<Record>
    {
        public string Key { get; set; }
        public string FileName { get; set; }
        public int Page { get; set; }

        public bool Equals(Record other)
        {
            if (other == null) return false;
            else if (other.Key == this.Key && other.FileName == this.FileName && other.Page == this.Page) return true;
            return false; 
        }
    }
}
