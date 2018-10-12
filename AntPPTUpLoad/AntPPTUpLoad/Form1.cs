using Newtonsoft.Json.Linq;
using SipoDataAcquisition;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AntPPTUpLoad
{
    public partial class Form1 : Form
    {
        public static string MainUrl = "http://www.antppt.com/fireman/login/index.html";
        public static string codeUrl = "http://www.antppt.com/fireman/login/getcode.html";
        public static string loginUrl = "http://www.antppt.com/fireman/login/index.html";
        public static CookieContainer cookie = new CookieContainer();
        public static string path = AppDomain.CurrentDomain.BaseDirectory;//程序运行根目录
        public static HttpHelper hh = null;
        public static string resPath = string.Empty;
        public static string sucessPath = string.Empty;
        public static string failPath = string.Empty;
        public static string accStr = string.Empty;
        public static string pwdStr = string.Empty;
        public static string codeStr = string.Empty;
        public static List<string> ResTxtPathList = new List<string>();
        public static int suc = 0;
        public static Task tTask = null;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            InitDefaultPath();
            GetYZM();
        }

        private void picClick(object sender, EventArgs e)
        {
            GetYZM();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            accStr = this.textBox1.Text;
            pwdStr = this.textBox2.Text;
            codeStr = this.textBox3.Text;
            resPath = this.textBox4.Text;
            sucessPath = this.textBox5.Text;
            failPath = this.textBox6.Text;
            if (string.IsNullOrEmpty(resPath) || string.IsNullOrEmpty(sucessPath) || string.IsNullOrEmpty(failPath))
            {
                MessageBox.Show("文件夹路径不能为空！", "提示");
                return;
            }
            if (string.IsNullOrEmpty(accStr) || string.IsNullOrEmpty(pwdStr) || string.IsNullOrEmpty(codeStr))
            {
                MessageBox.Show("账号密码与验证码不能为空！", "提示");
                return;
            }

            CreateDefaultPath(resPath, sucessPath, failPath);

            if (ToLogin(accStr, pwdStr, codeStr))
            {
                GetResTxtPath();
                Form2 f2 = new Form2();
                this.Hide();
                f2.Show();
            }

        }
        public void F2StartF1()
        {
            Task task = Task.Factory.StartNew(StartThreadToWatch);
        }
        /// <summary>
        /// 开启线程去监视任务
        /// </summary>
        public void StartThreadToWatch()
        {
            while (true)
            {
                suc = 0;
                if (GetResTxtPath().Count > 0)
                {
                    tTask = Task.Factory.StartNew(StartUpLoad);
                    tTask.Wait();
                }
                Thread.Sleep(1000 * 60 * 1);
            }
        }
        public void StartUpLoad()
        {
            if (ResTxtPathList.Count > 0)
            {
                for (int i = 0; i < ResTxtPathList.Count; i++)
                {
                    TxtSetting txtObj = ReadTxtToObj(ResTxtPathList[i]);
                    ToUpLoadRes(txtObj);
                    suc++;
                }
            }
        }
        private void Resource_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.textBox4.Text = fbd.SelectedPath;//选定目录
            }
        }
        private void Success_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.textBox5.Text = fbd.SelectedPath;//选定目录
            }
        }
        private void Fail_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            if (fbd.ShowDialog() == DialogResult.OK)
            {
                this.textBox6.Text = fbd.SelectedPath;//选定目录
            }
        }
        /// <summary>
        /// 获取验证码
        /// </summary>
        public void GetYZM()
        {
            Bitmap bitYzm = null;
            hh = new HttpHelper("183.30.204.174", 9999);
            hh.GetHtmlData(MainUrl, cookie);//获取cookie
            byte[] bytelist = hh.DowloadCheckImg(codeUrl, cookie, path);//下载验证码
            if (bytelist.Length > 0)
            {
                MemoryStream ms1 = new MemoryStream(bytelist);
                bitYzm = (Bitmap)Image.FromStream(ms1);
                ms1.Close();
                this.pictureBox1.Image = bitYzm;
            }
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml("htmlStr");
        }
        /// <summary>
        /// 登录
        /// </summary>
        public bool ToLogin(string accStr, string pwdStr, string codeStr)
        {
            string postData = string.Format("account={0}&password={1}&code={2}", accStr, pwdStr, codeStr);
            string htmlStr = hh.PostWebRequest(loginUrl, postData, cookie);
            if (htmlStr.Contains("登录成功"))
                return true;
            MessageBox.Show("登录失败！", "提示");
            GetYZM();
            return false;
        }
        /// <summary>
        /// 上传
        /// </summary>
        /// <param name="txtObj"></param>
        public void ToUpLoadRes(TxtSetting txtObj)
        {
            DateTime dt = DateTime.Now;
            string m, d = string.Empty;
            if (dt.Month < 10)
                m = "0" + dt.Month.ToString();
            else
                m = dt.Month.ToString();
            if (dt.Day < 10)
                d = "0" + dt.Day.ToString();
            else
                d = dt.Day.ToString();
            string timeFolder = dt.Year.ToString() + m + d;
            List<string> uploadPath = new List<string>();
            string title = txtObj.title;
            string policy = @"eyJleHBpcmF0aW9uIjoiMjAyMC0wMS0wMVQxMjowMDowMC4wMDBaIiwiY29uZGl0aW9ucyI6W1siY29udGVudC1sZW5ndGgtcmFuZ2UiLDAsMTA0ODU3NjAwMF1dfQ==";
            string oSSAccessKeyId = "LTAIFwHQY1lUbk1D";
            string success_action_status = "200";
            string signature = @"oNFWy4ePoF8sxRRf5ylB/ihJ7+M=";
            string[] extArr = { ".JPG", ".swf", ".zip", ".JPG" };
            string[] contentTypeArr = { "image/jpeg", "application/x-shockwave-flash", "application/x-zip-compressed", "image/jpeg" };
            string[] PathArr = { txtObj.smallPicPath, txtObj.videoPath, txtObj.zipPath, txtObj.lognPicPath };
            string ossUrl = "http://antppt.oss-cn-beijing.aliyuncs.com/";
            string outResult = "";

            for (int i = 0; i < extArr.Length; i++)
            {
                HttpUpLoad hu = null;
                hu = new HttpUpLoad();
                string encryption = MD5(title);
                uploadPath.Add(encryption);
                hu.SetFieldValue("name", title + extArr[i]);
                hu.SetFieldValue("key", timeFolder + "/" + encryption + extArr[i]);
                hu.SetFieldValue("policy", policy);
                hu.SetFieldValue("OSSAccessKeyId", oSSAccessKeyId);
                hu.SetFieldValue("success_action_status", success_action_status);
                hu.SetFieldValue("signature", signature);
                hu.SetFieldValue("file", title + extArr[i], contentTypeArr[i], GetFileData(PathArr[i]));
                hu.Upload(ossUrl, cookie, out outResult);
            }

            HttpUpLoad hl = new HttpUpLoad();
            hl.SetFieldValue("menu_id", txtObj.menu_id);
            string[] sArr = txtObj.style_ids.Split(',');
            for (int i = 0; i < sArr.Length; i++)
            {
                if (!string.IsNullOrEmpty(sArr[i]))
                    hl.SetFieldValue("style_ids[]", sArr[i]);
            }

            hl.SetFieldValue("title", txtObj.title);
            hl.SetFieldValue("keyword", txtObj.keyword);
            hl.SetFieldValue("related_search", txtObj.related_search);
            hl.SetFieldValue("image", ossUrl + timeFolder + "/" + uploadPath[0] + extArr[0]);
            hl.SetFieldValue("video", ossUrl + timeFolder + "/" + uploadPath[1] + extArr[1]);
            hl.SetFieldValue("txt_file", ossUrl + timeFolder + "/" + uploadPath[2] + extArr[2]);
            hl.SetFieldValue("file_size", GetFileSize(txtObj.zipPath));
            hl.SetFieldValue("price", txtObj.price);
            hl.SetFieldValue("software", txtObj.soft);
            hl.SetFieldValue("page", txtObj.page);
            hl.SetFieldValue("proportion", txtObj.proportion);
            hl.SetFieldValue("score", txtObj.score);
            hl.SetFieldValue("free_status", txtObj.free_status);
            hl.SetFieldValue("is_tuijian", txtObj.is_tuijian);
            hl.SetFieldValue("notes", txtObj.notes);
            hl.SetFieldValue("complete", txtObj.complete);
            hl.SetFieldValue("sort", txtObj.sort);
            hl.SetFieldValue("image_list", "");
            hl.SetFieldValue("pictureurls[]", ossUrl + timeFolder + "/" + uploadPath[3] + extArr[3]);
            hl.SetFieldValue("describe", txtObj.describe);
            hl.SetFieldValue("source_describe", txtObj.source_describe);

            bool res = hl.PostData("http://www.antppt.com/fireman/content/add.html", cookie);
            string parentFolder = resPath + @"\" + txtObj.title;

            if (res)
            {
                if (Directory.Exists(parentFolder))
                {
                    DirectoryInfo pfDir = new DirectoryInfo(parentFolder);
                    foreach (FileSystemInfo fsi in pfDir.GetFileSystemInfos())
                    {
                        if (fsi is FileInfo)
                        {
                            if (!Directory.Exists(sucessPath + @"\" + txtObj.title))
                                Directory.CreateDirectory(sucessPath + @"\" + txtObj.title);
                            File.Move(fsi.FullName, sucessPath + @"\" + txtObj.title + @"\" + fsi.Name);
                        }
                    }
                    pfDir.Delete();
                }
            }
            else
            {
                if (Directory.Exists(parentFolder))
                {
                    DirectoryInfo pfDir = new DirectoryInfo(parentFolder);
                    foreach (FileSystemInfo fsi in pfDir.GetFileSystemInfos())
                    {
                        if (fsi is FileInfo)
                        {
                            if (!Directory.Exists(failPath + @"\" + txtObj.title))
                                Directory.CreateDirectory(failPath + @"\" + txtObj.title);
                            File.Move(fsi.FullName, failPath + @"\" + txtObj.title + @"\" + fsi.Name);
                        }
                    }
                    pfDir.Delete();
                }
            }
        }
        /// <summary> /// 加密字符串   
        /// </summary>  
        /// <param name="str">要加密的字符串</param>  
        /// <param name="encryptKey">加密密钥</param>  
        /// <returns>加密后的字符串</returns>  
        static string MD5(string str, int code = 16)
        {
            Thread.Sleep(1000);
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            string timeSpan = Convert.ToInt64(ts.TotalSeconds).ToString();
            str += timeSpan;
            if (code == 16) //16位MD5加密（取32位加密的9~25字符） 
            {
                return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(str, "MD5").ToLower().Substring(8, 16);
            }
            else//32位加密 
            {
                return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(str, "MD5").ToLower();
            }
        }
        /// <summary>
        /// 将文件转换成byte[] 数组
        /// </summary>
        /// <param name="fileUrl">文件路径文件名称</param>
        /// <returns>byte[]</returns>
        protected byte[] GetFileData(string fileUrl)
        {
            FileStream fs = new FileStream(fileUrl, FileMode.Open, FileAccess.Read);
            try
            {
                byte[] buffur = new byte[fs.Length];
                fs.Read(buffur, 0, (int)fs.Length);

                return buffur;
            }
            catch (Exception ex)
            {
                return null;
            }
            finally
            {
                if (fs != null)
                {
                    //关闭资源
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// 将字节数组转换成指定类型的文件
        /// </summary>
        /// <param name="byteArr">字节数组</param>
        /// <param name="fileFullName">要生成的文件名（完整路径，包含文件名与文件类型）</param>
        public void SaveByteToFile(byte[] byteArr, string fileFullName)
        {
            try
            {
                using (MemoryStream ms = new MemoryStream(byteArr))
                {
                    using (FileStream fs = new FileStream(fileFullName, FileMode.OpenOrCreate))
                    {
                        ms.WriteTo(fs);
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        /// <summary>
        /// 读取目录下的文件
        /// </summary>
        /// <param name="extNameStr">扩展名</param>
        public List<string> GetFiles(string fileDir, string extNameStr)
        {
            List<string> pathList = new List<string>();
            if (!string.IsNullOrEmpty(fileDir))
            {
                if (!Directory.Exists(fileDir))
                    return null;
                DirectoryInfo mydir = new DirectoryInfo(fileDir);
                foreach (FileSystemInfo fsi in mydir.GetFileSystemInfos())
                {
                    if (fsi is FileInfo)
                    {
                        string extName = Path.GetExtension(fsi.FullName); //获取扩展名  
                        if (extName.ToUpper() == extNameStr.ToUpper())
                        {
                            pathList.Add(fsi.FullName);
                        }
                    }
                }
            }
            return pathList;
        }
        /// <summary>
        /// 获取资源配置文件
        /// </summary>
        public List<string> GetResTxtPath()
        {
            ResTxtPathList.Clear();
            List<string> tempList = new List<string>();
            DirectoryInfo dir = new DirectoryInfo(resPath);
            List<DirectoryInfo> ResFolederList = dir.GetDirectories().ToList();
            foreach (var item in ResFolederList)
            {
                dir = new DirectoryInfo(item.FullName);
                int qe = dir.GetFiles().Length;
                if (dir.GetFiles().Length < 4)
                    continue;
                foreach (FileSystemInfo fsi in dir.GetFileSystemInfos())
                {
                    if (fsi is FileInfo)
                    {
                        string extName = Path.GetExtension(fsi.FullName); //获取扩展名  
                        if (extName.ToUpper() == ".TXT")
                        {
                            ResTxtPathList.Add(fsi.FullName);
                        }
                    }
                }
            }
            tempList = ResTxtPathList;
            return tempList;
        }

        public TxtSetting ReadTxtToObj(string txtPath)
        {
            string tempMenu_id = "", tempStyle_ids = "", tempTitle = "", tempKeyword = "",
                tempRelated_search = "", tempSmallPicPath = "", tempVideoPath = "", tempZipPath = "",
                tempLognPicPath = "", tempPrice = "", tempSoft = "", tempPage = "", tempProportion = "",
                tempScore = "", tempFree_status = "", tempIs_tuijian = "", tempNotes = "",
                tempComplete = "", tempSort = "", tempDescribe = "", tempSource_describe = "";

            string txtStr = File.ReadAllText(txtPath, Encoding.Default);
            string[] txtArr = SplitByString(txtStr, "\r\n");
            for (int i = 0; i < txtArr.Length; i++)
            {
                string[] txtKv = SplitByString(txtArr[i], "->");
                switch (txtKv[0])
                {
                    case "menu_id":
                        {
                            tempMenu_id = txtKv[1];
                            break;
                        }
                    case "style_ids":
                        {
                            tempStyle_ids = txtKv[1];
                            break;
                        }
                    case "title":
                        {
                            tempTitle = txtKv[1];
                            break;
                        }
                    case "keyword":
                        {
                            tempKeyword = txtKv[1];
                            break;
                        }
                    case "related_search":
                        {
                            tempRelated_search = txtKv[1];
                            break;
                        }
                    case "smallPicPath":
                        {
                            tempSmallPicPath = txtKv[1];
                            break;
                        }
                    case "videoPath":
                        {
                            tempVideoPath = txtKv[1];
                            break;
                        }
                    case "zipPath":
                        {
                            tempZipPath = txtKv[1];
                            break;
                        }
                    case "lognPicPath":
                        {
                            tempLognPicPath = txtKv[1];
                            break;
                        }
                    case "price":
                        {
                            tempPrice = txtKv[1];
                            break;
                        }
                    case "soft":
                        {
                            tempSoft = txtKv[1];
                            break;
                        }
                    case "page":
                        {
                            tempPage = txtKv[1];
                            break;
                        }
                    case "proportion":
                        {
                            tempProportion = txtKv[1];
                            break;
                        }
                    case "score":
                        {
                            tempScore = txtKv[1];
                            break;
                        }
                    case "free_status":
                        {
                            tempFree_status = txtKv[1];
                            break;
                        }
                    case "is_tuijian":
                        {
                            tempIs_tuijian = txtKv[1];
                            break;
                        }
                    case "notes":
                        {
                            tempNotes = txtKv[1];
                            break;
                        }
                    case "complete":
                        {
                            tempComplete = txtKv[1];
                            break;
                        }
                    case "sort":
                        {
                            tempSort = txtKv[1];
                            break;
                        }
                    case "describe":
                        {
                            tempDescribe = txtKv[1];
                            break;
                        }
                    case "source_describe":
                        {
                            tempSource_describe = txtKv[1];
                            break;
                        }
                    default:
                        break;
                }
            }
            return new TxtSetting
            {
                menu_id = tempMenu_id,
                style_ids = tempStyle_ids,
                title = tempTitle,
                keyword = tempKeyword,
                related_search = tempRelated_search,
                smallPicPath = tempSmallPicPath,
                videoPath = tempVideoPath,
                zipPath = tempZipPath,
                lognPicPath = tempLognPicPath,
                price = tempPrice,
                soft = tempSoft,
                page = tempPage,
                proportion = tempProportion,
                score = tempScore,
                free_status = tempFree_status,
                is_tuijian = tempIs_tuijian,
                notes = tempNotes,
                complete = tempComplete,
                sort = tempSort,
                describe = tempDescribe,
                source_describe = tempSource_describe
            };
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
        /// <summary>
        /// 计算文件大小函数(保留两位小数),Size为字节大小
        /// </summary>
        /// <param name="path">初始文件大小</param>
        /// <returns></returns>
        public static string GetFileSize(string path)
        {
            FileInfo fileInfo = new FileInfo(path);
            long size = fileInfo.Length;//初始文件大小,字节长度

            var num = 1024.00; //byte

            if (size < num)
                return size + "B";
            if (size < Math.Pow(num, 2))
                return (size / num).ToString("f2") + "K"; //kb
            if (size < Math.Pow(num, 3))
                return (size / Math.Pow(num, 2)).ToString("f2") + "M"; //M
            if (size < Math.Pow(num, 4))
                return (size / Math.Pow(num, 3)).ToString("f2") + "G"; //G

            return (size / Math.Pow(num, 4)).ToString("f2") + "T"; //T
        }
        /// <summary>
        /// 保存默认路径
        /// </summary>
        /// <param name="p1"></param>
        /// <param name="p2"></param>
        /// <param name="p3"></param>
        public void CreateDefaultPath(string p1, string p2, string p3)
        {
            string defaultSetting = path + "DefaultSetting.ini";
            string content = string.Format("Path1=>{0}\r\nPath2=>{1}\r\nPath3=>{2}\r\n", p1, p2, p3);
            File.WriteAllText(defaultSetting, content);
        }
        /// <summary>
        /// 初始化默认路径
        /// </summary>
        public void InitDefaultPath()
        {
            string defaultSetting = path + "DefaultSetting.ini";
            if (File.Exists(defaultSetting))
            {
                string contStr = File.ReadAllText(defaultSetting);
                string[] pathArr = SplitByString(contStr, "\r\n");
                this.textBox4.Text = SplitByString(pathArr[0], "=>")[1];
                this.textBox5.Text = SplitByString(pathArr[1], "=>")[1];
                this.textBox6.Text = SplitByString(pathArr[2], "=>")[1];
            }
        }
    }
}
