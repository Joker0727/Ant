using MyPPT;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AntPPT
{
    public partial class Form1 : Form
    {
        public string baseDir = System.AppDomain.CurrentDomain.BaseDirectory;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitSetting();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog pptFolder = new FolderBrowserDialog();

            if (pptFolder.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = pptFolder.SelectedPath;//选定目录
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog targetFolder = new FolderBrowserDialog();

            if (targetFolder.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = targetFolder.SelectedPath;//选定目录
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string pptF = this.textBox1.Text;
            string pptT = this.textBox2.Text;
            string wordStr = this.textBox3.Text;
            string spaceStr = this.textBox6.Text;
            string totalWidthStr = this.textBox7.Text;
            string errorPath = this.textBox4.Text;
            string successPath = this.textBox5.Text;
            double scale = double.Parse(this.textBox8.Text);


            if (string.IsNullOrEmpty(successPath) || string.IsNullOrEmpty(errorPath) || string.IsNullOrEmpty(pptF) || string.IsNullOrEmpty(pptT) || string.IsNullOrEmpty(spaceStr) || string.IsNullOrEmpty(totalWidthStr))
            {
                MessageBox.Show("目录不能为空！", "提示");
                return;
            }

            ResetSetting(pptF, pptT, errorPath, successPath, wordStr);

            int space = int.Parse(spaceStr);
            int totalWidth = int.Parse(totalWidthStr);

            StartDealPPT sdp = new StartDealPPT(pptF, pptT, errorPath, successPath, wordStr, space, totalWidth, scale);

            Thread th = new Thread(sdp.StartDeal);
            th.IsBackground = true;
            th.ApartmentState = ApartmentState.STA;
            th.Start();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog targetFolder = new FolderBrowserDialog();

            if (targetFolder.ShowDialog() == DialogResult.OK)
            {
                this.textBox4.Text = targetFolder.SelectedPath;//选定目录
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog targetFolder = new FolderBrowserDialog();

            if (targetFolder.ShowDialog() == DialogResult.OK)
            {
                this.textBox5.Text = targetFolder.SelectedPath;//选定目录
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("你确定要关闭吗！", "提示信息", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (result == DialogResult.OK)
            {
                PPTHelper ppt = new PPTHelper();
                ppt.PPTClose();
                e.Cancel = false;  //点击OK   
            }
            else
            {
                e.Cancel = true;
            }
        }
        /// <summary>
        /// 初始化设置
        /// </summary>
        public void InitSetting()
        {
            string pathConfig = baseDir + "DefaultPath.ini";
            string wordsConfig = baseDir + "DefaultWords.ini";

            if (File.Exists(wordsConfig))
            {
                string defaultPathStr = File.ReadAllText(wordsConfig, Encoding.Default);
                if (defaultPathStr.Length > 0)
                {
                    this.textBox3.Text = defaultPathStr;
                }
            }

            if (File.Exists(pathConfig))
            {
                string defaultPathStr = File.ReadAllText(pathConfig, Encoding.Default);
                if (defaultPathStr.Length > 0)
                {
                    string[] defaultPathKVStr = Regex.Split(defaultPathStr, "\r\n", RegexOptions.IgnoreCase);
                    for (int i = 0; i < defaultPathKVStr.Length; i++)
                    {
                        if (i == 4)
                            continue;
                        string[] defaultPathKV = defaultPathKVStr[i].Split('=');
                        if (defaultPathKV.Length > 0)
                        {
                            if (!string.IsNullOrEmpty(defaultPathKV[1]))
                            {
                                switch (defaultPathKV[0])
                                {
                                    case "SourcePath":
                                        {
                                            this.textBox1.Text = defaultPathKV[1];
                                            break;
                                        }
                                    case "OutPath":
                                        {
                                            this.textBox2.Text = defaultPathKV[1];
                                            break;
                                        }
                                    case "ExceptionPath":
                                        {
                                            this.textBox4.Text = defaultPathKV[1];
                                            break;
                                        }
                                    case "Success":
                                        {
                                            this.textBox5.Text = defaultPathKV[1];
                                            break;
                                        }
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// 重置默认设置
        /// </summary>
        public void ResetSetting(string pptF, string pptT, string errorPath, string successPath, string wordStr)
        {
            string pathConfig = baseDir + "DefaultPath.ini";
            string wordsConfig = baseDir + "DefaultWords.ini";

            if (File.Exists(wordsConfig))
                File.Delete(wordsConfig);
            File.WriteAllText(wordsConfig, wordStr, Encoding.Default);

            if (File.Exists(pathConfig))
                File.Delete(pathConfig);

            string path1 = "SourcePath=" + pptF;
            string path2 = "OutPath=" + pptT;
            string path3 = "ExceptionPath=" + errorPath;
            string path4 = "Success=" + successPath;
            string[] pathArr = { path1, path2, path3, path4 };

            File.WriteAllLines(pathConfig, pathArr, Encoding.Default);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            Form2 f2 = new Form2();
            //f2.Show();
            f2.ShowDialog();
        }
    }
}
