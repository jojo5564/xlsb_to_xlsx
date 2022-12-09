using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
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
        private string OutputPath = "";
        private readonly IniManager iniManager = new IniManager(Application.StartupPath + "\\Config.ini");

        public Form1()
        {
            InitializeComponent();

            textBox1.Text = iniManager.ReadIniFile("Location", "Output", " ");
            OutputPath = iniManager.ReadIniFile("Location", "Output", " ");
        }
        
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            listBox1.Items.Clear();
            if (OutputPath == "")
            {
                MessageBox.Show("Error : Output location is null");
                button1.Enabled = true;
                button2.Enabled = true;
                return;
            }
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Title = "Please select .xlsb files";
            dialog.Filter = "*.xlsb|*.xlsb";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                Thread t1 = new Thread(new ParameterizedThreadStart(Convert));
                t1.Start(dialog.FileNames);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var tmp = GetSavePath();
            if (tmp != "")
            {
                OutputPath = tmp;
                textBox1.Text = tmp;
                iniManager.WriteIniFile("Location", "Output", tmp);
            }
        }

        private string GetSavePath()
        {
            FolderBrowserDialog savedialog = new FolderBrowserDialog();
            savedialog.ShowDialog();
            return savedialog.SelectedPath;
        }

        private void Convert(object obj)
        {
            string[] files = (string[])obj;
            ListBoxControl(listBox1, "Add", "Total : " + files.Length + " files");
            for (int i = 0; i < files.Length; i++)
            {
                ListBoxControl(listBox1, "Add", String.Format("Converting {0}/{1} : {2} ...", i + 1, files.Length, files[i].Split('\\').Last()));
                
                try
                {
                    XlsbToXlsx(files[i]);
                }
                catch (Exception ex)
                {
                    ListBoxControl(listBox1, "Add", String.Format("fail {0}/{1} : {2} ...", i + 1, files.Length, files[i].Split('\\').Last()));
                }
            }
            ListBoxControl(listBox1, "Add", "succeed");
            ButtonControl(button1, true);
            ButtonControl(button2, true);
        }

        private void XlsbToXlsx(object obj)
        {
            string file = (String)obj;
            var workbook = new Workbook(file);
            string newFileName = (file.Split(new[] { ".xlsb" }, StringSplitOptions.None)[0] + ".xlsx").Split('\\').Last();
            workbook.Save(OutputPath + '\\' + newFileName);
        }

        #region delegate
        private delegate void ListBoxControlCallback(ListBox listBox, string command, string str);
        public delegate void ButtonControlCallback(Button button, bool enable);

        public void ListBoxControl(ListBox listBox, string command, string str)
        {
            if (listBox.InvokeRequired)
            {
                ListBoxControlCallback d = new ListBoxControlCallback(ListBoxControl);
                this.Invoke(d, new object[] { listBox, command, str });
            }
            else
            {
                if (command == "Add")
                {
                    listBox.Items.Add(str);
                    listBox.TopIndex = listBox.Items.Count - 1;
                }
                else if (command == "Clear")
                {
                    listBox.Items.Clear();
                }
            }
        }
        public void ButtonControl(Button button, bool enable)
        {
            if (button.InvokeRequired)
            {
                ButtonControlCallback d = new ButtonControlCallback(ButtonControl);
                this.Invoke(d, new object[] { button, enable });
            }
            else
            {
                button.Enabled = enable;
            }
        }
        #endregion

    }
}
