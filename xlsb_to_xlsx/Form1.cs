using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;

namespace xlsb_to_xlsx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Title = "Please select .xlsb files";
            dialog.Filter = "*.xlsb|*.xlsb";
            string[] files;

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                files = dialog.FileNames;
                listBox1.Items.Add("Total : " + files.Length + " files");
                foreach (var file in files)
                {
                    listBox1.Items.Add("Converting : " + file.Split('\\').Last());
                    try
                    {
                        Thread t1 = new Thread(new ParameterizedThreadStart(Convert));
                        t1.Start(file);
                        while (t1.ThreadState == ThreadState.Running)
                        {
                            this.Update();
                            Thread.Sleep(100);
                        }
                    }
                    catch (Exception ex)
                    {
                        listBox1.Items.Add("Failed");
                    }
                }
                listBox1.Items.Add("succeed");
            }
        }

        private void Convert(object objs)
        {
            string file = (String)objs;
            var workbook = new Workbook(file);
            string newFileName = file.Split(new[] { ".xlsb" }, StringSplitOptions.None)[0] + ".xlsx";
            workbook.Save(newFileName);
        }
    }
}
