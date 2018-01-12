using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Word;

namespace FileNames
{
    public partial class Form1 : Form
    {

        public static string docfilepath;
        public static string[] files;

        public Form1()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        { 
            this.folderBrowserDialog1 = new FolderBrowserDialog();

            if (folderBrowserDialog1.ShowDialog(this) == DialogResult.OK)
            {

                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath);
            }
            textBox1.Text = folderBrowserDialog1.SelectedPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.progressBar1.Value = 0;

            this.timer1.Interval = 100;
            this.timer1.Enabled = true;

            try
            {
                for (int i = 0; i < files.Count(); i++)
                {
                    OpenWordDocument(files[i]);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void OpenWordDocument(string path)
        {
            
            object objMissing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word._Document objDoc;

            Microsoft.Office.Interop.Word._Application objWord = new Microsoft.Office.Interop.Word.Application();

            object fileName = path;
            objDoc = objWord.Documents.Open(ref fileName,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing);

            var sourcefile = new FileInfo(path);
            object newFileName = sourcefile.FullName.Replace(".doc", ".docx");
            object docFormat = WdSaveFormat.wdFormatXMLDocument;

            objDoc.SaveAs(ref newFileName, ref docFormat,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing, ref objMissing,
                ref objMissing, ref objMissing, ref objMissing, ref objMissing);
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (this.progressBar1.Value < 100)
            {
                this.progressBar1.Value++;
                if (this.progressBar1.Value == 100)
                {
                    MessageBox.Show("Files are Successfully Converted!");
                }
            }
            else
            {
                this.timer1.Enabled = false;
            } 
        }

    }
}
