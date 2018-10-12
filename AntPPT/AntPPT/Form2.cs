using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AntPPT
{
    public partial class Form2 : Form
    {
        public string baseDir = System.AppDomain.CurrentDomain.BaseDirectory;
        public string defaultColumPath = System.AppDomain.CurrentDomain.BaseDirectory + "columSetting.ini";
        public string defaultStylePath = System.AppDomain.CurrentDomain.BaseDirectory + "styleSetting.ini";
        public string defaultKeyWordPath = System.AppDomain.CurrentDomain.BaseDirectory + "keyWordSetting.ini";
        public string defaultSearchPath = System.AppDomain.CurrentDomain.BaseDirectory + "searchSetting.ini";
        public string defaultSettingPath = System.AppDomain.CurrentDomain.BaseDirectory + "DefaultSetting.ini";

        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            InitDefaultSetting();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (SaveDefaultSetting())
            {
                MessageBox.Show("默认设置保存成功！", "提示");
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Colum_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();//首先，实例化对话框类实例
            openDialog.Filter = "文本文件|*.txt";
            if (DialogResult.OK == openDialog.ShowDialog())//然后，判断如果当前用户在对话框里点击的是OK按钮的话。
            {
                string filename = openDialog.FileName; //将打开文件对话框的FileName属性传递到你的字符串进行处理
                this.textBox1.Text = filename;
            }
        }

        private void Style_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();//首先，实例化对话框类实例
            openDialog.Filter = "文本文件|*.txt";
            if (DialogResult.OK == openDialog.ShowDialog())//然后，判断如果当前用户在对话框里点击的是OK按钮的话。
            {
                string filename = openDialog.FileName; //将打开文件对话框的FileName属性传递到你的字符串进行处理
                this.textBox2.Text = filename;
            }
        }

        private void KeyWords_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();//首先，实例化对话框类实例
            openDialog.Filter = "文本文件|*.txt";
            if (DialogResult.OK == openDialog.ShowDialog())//然后，判断如果当前用户在对话框里点击的是OK按钮的话。
            {
                string filename = openDialog.FileName; //将打开文件对话框的FileName属性传递到你的字符串进行处理
                this.textBox4.Text = filename;
            }
        }

        private void Search_Click(object sender, EventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();//首先，实例化对话框类实例
            openDialog.Filter = "文本文件|*.txt";
            if (DialogResult.OK == openDialog.ShowDialog())//然后，判断如果当前用户在对话框里点击的是OK按钮的话。
            {
                string filename = openDialog.FileName; //将打开文件对话框的FileName属性传递到你的字符串进行处理
                this.textBox5.Text = filename;
            }
        }
        /// <summary>
        /// 保存默认设置
        /// </summary>
        public bool SaveDefaultSetting()
        {
            string columPath = this.textBox1.Text;
            string stylePath = this.textBox2.Text;
            string keyWordPath = this.textBox4.Text;
            string searchPath = this.textBox5.Text;
            if (string.IsNullOrEmpty(columPath) || string.IsNullOrEmpty(stylePath) ||
                string.IsNullOrEmpty(keyWordPath) || string.IsNullOrEmpty(searchPath))
            {
                MessageBox.Show("默认设置路径不能为空！", "提示");
                return false;
            }

            if (File.Exists(columPath))
            {
                if (File.Exists(defaultColumPath))
                    File.Delete(defaultColumPath);
                File.Copy(columPath, defaultColumPath);
            }
            else
            {
                MessageBox.Show("栏目配置文件不存在！", "提示");
                return false;
            }
            if (File.Exists(stylePath))
            {
                if (File.Exists(defaultStylePath))
                    File.Delete(defaultStylePath);
                File.Copy(stylePath, defaultStylePath);
            }
            else
            {
                MessageBox.Show("风格配置文件不存在！", "提示");
                return false;
            }
            if (File.Exists(keyWordPath))
            {
                if (File.Exists(defaultKeyWordPath))
                    File.Delete(defaultKeyWordPath);
                File.Copy(keyWordPath, defaultKeyWordPath);
            }
            else
            {
                MessageBox.Show("关键词配置文件不存在！", "提示");
                return false;
            }
            if (File.Exists(searchPath))
            {
                if (File.Exists(defaultSearchPath))
                    File.Delete(defaultSearchPath);
                File.Copy(searchPath, defaultSearchPath);
            }
            else
            {
                MessageBox.Show("搜索词配置文件不存在！", "提示");
                return false;
            }

            string price = this.textBox6.Text;
            string software = this.textBox7.Text;
            string proportion = this.textBox8.Text;
            string score = this.textBox9.Text;
            string isfree = this.textBox10.Text;
            string istuijian = this.textBox11.Text;
            string notes = this.textBox12.Text;
            string complete = this.textBox13.Text;
            string sort = this.textBox3.Text;

            if (string.IsNullOrEmpty(price) || string.IsNullOrEmpty(software) || string.IsNullOrEmpty(proportion) ||
                string.IsNullOrEmpty(score) || string.IsNullOrEmpty(isfree) || string.IsNullOrEmpty(istuijian) ||
                string.IsNullOrEmpty(notes) || string.IsNullOrEmpty(complete) || string.IsNullOrEmpty(sort))
            {
                MessageBox.Show("请补全默认设置！", "提示");
                return false;
            }

            string defaultStr = @"colum->" + defaultColumPath + "\r\nstyle->" + defaultStylePath + "\r\nkeyword->" + defaultKeyWordPath
                + "\r\nsearch->" + defaultSearchPath + "\r\nprice->" + price + "\r\nsoftware->" + software
                + "\r\nproportion->" + proportion + "\r\nscore->" + score + "\r\nisfree->" + isfree
                + "\r\nistuijian->" + istuijian + "\r\nnotes->" + notes + "\r\ncomplete->" + complete
                + "\r\nsort->" + sort;

            File.WriteAllText(defaultSettingPath, defaultStr);
            return true;
        }
        /// <summary>
        /// 初始化默认设置
        /// </summary>
        /// <returns></returns>
        public bool InitDefaultSetting()
        {
            if (!File.Exists(defaultSettingPath))
            {
                MessageBox.Show("默认设置配置文件不存在，请重新设置！", "提示");
                return false;
            }
            string defaultStr = File.ReadAllText(defaultSettingPath);
            string[] dStrArr = SplitByString(defaultStr, "\r\n");
            for (int i = 0; i < dStrArr.Length; i++)
            {
                if (!string.IsNullOrEmpty(dStrArr[i]))
                {
                    string[] dKV = SplitByString(dStrArr[i], "->");
                    if (dKV.Length == 2)
                    {
                        switch (dKV[0])
                        {
                            case "colum":
                                {
                                    this.textBox1.Text = dKV[1];
                                    break;
                                }
                            case "style":
                                {
                                    this.textBox2.Text = dKV[1];
                                    break;
                                }
                            case "keyword":
                                {
                                    this.textBox4.Text = dKV[1];
                                    break;
                                }
                            case "search":
                                {
                                    this.textBox5.Text = dKV[1];
                                    break;
                                }
                            case "price":
                                {
                                    this.textBox6.Text = dKV[1];
                                    break;
                                }
                            case "software":
                                {
                                    this.textBox7.Text = dKV[1];
                                    break;
                                }
                            case "proportion":
                                {
                                    this.textBox8.Text = dKV[1];
                                    break;
                                }
                            case "score":
                                {
                                    this.textBox9.Text = dKV[1];
                                    break;
                                }
                            case "isfree":
                                {
                                    this.textBox10.Text = dKV[1];
                                    break;
                                }
                            case "istuijian":
                                {
                                    this.textBox11.Text = dKV[1];
                                    break;
                                }
                            case "notes":
                                {
                                    this.textBox12.Text = dKV[1];
                                    break;
                                }
                            case "complete":
                                {
                                    this.textBox13.Text = dKV[1];
                                    break;
                                }
                            case "sort":
                                {
                                    this.textBox3.Text = dKV[1];
                                    break;
                                }
                            default:
                                break;
                        }
                    }
                }
            }

            return true;
        }
        /// <summary>
        /// 用字符串来分割成数组
        /// </summary>
        /// <param name="originalString">原字符串</param>
        /// <param name="strKey">以strKey作为分割符</param>
        /// <returns></returns>
        public string[] SplitByString(string originalString, string strKey)
        {
            string[] sArray = Regex.Split(originalString, strKey, RegexOptions.IgnoreCase);
            return sArray;
        }

    }
}
