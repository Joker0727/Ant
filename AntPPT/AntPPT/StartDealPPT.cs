using GTR;
using Microsoft.Office.Interop.PowerPoint;
using MyPPT;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static Compress.ZipUtility;

namespace AntPPT
{
    public class StartDealPPT
    {
        public List<string> pptPathList = new List<string>();
        public List<string> picPathList = new List<string>();
        public string filePath = string.Empty;
        public string targetPath = string.Empty;
        public PPTHelper ppt = new PPTHelper();
        public string wordStr = string.Empty;
        public List<KeyValuePair<string, string>> oldNewList = new List<KeyValuePair<string, string>>();
        public string baseDir = System.AppDomain.CurrentDomain.BaseDirectory;
        public int space, totalWidth = 0;
        public string fileNameWithoutExtension = string.Empty;
        public string fileNameWithExtension = string.Empty;
        public string singlePath = string.Empty;
        public string errorPath = string.Empty;
        public string successPath = string.Empty;
        public string defaultSettingPath = System.AppDomain.CurrentDomain.BaseDirectory + "DefaultSetting.ini";
        public string keyWordTemp = string.Empty;
        public string pptClassName = "PPTFrameClass";
        public double scale = 0;

        public StartDealPPT()
        {

        }
        public StartDealPPT(string pptF, string pptT, string errorPath, string successPath, string wordStr, int space, int totalWidth, double xyscale=1)
        {
            this.filePath = pptF;
            this.targetPath = pptT;
            this.wordStr = wordStr;
            this.space = space;
            this.totalWidth = totalWidth;
            this.errorPath = errorPath;
            this.successPath = successPath;
            this.scale = xyscale;

            GetFiles(filePath, ".pptx");//读取ppt路径
            StrToKeyValue();//转换文字替换字符串键值对
        }

        public void StartDeal()
        {
            if (pptPathList.Count <= 0)
                return;

            foreach (var path in pptPathList)
            {
                fileNameWithoutExtension = Path.GetFileNameWithoutExtension(path);//获取文件名，没有扩展名的文件名 
                fileNameWithExtension = Path.GetFileName(path);//获取文件名
                singlePath = targetPath + @"\" + fileNameWithoutExtension + @"\";
                if (Directory.Exists(singlePath))
                {
                    DirectoryInfo subdir = new DirectoryInfo(singlePath);
                    subdir.Delete(true);
                }
                ppt.PPTOpen(path);//打开PPT
                if (ppt.objApp == null || ppt.objPresSet == null)
                {
                    string tarFName = Path.GetFileName(path);
                    string errorFullName = errorPath + @"\" + tarFName;
                    if (File.Exists(path))
                    {
                        if (!Directory.Exists(errorPath))
                            Directory.CreateDirectory(errorPath);
                        File.Move(path, errorFullName);
                    }
                    continue;
                }
                if (IsOnlyRead(path))
                {
                    try
                    {
                        string tarFName = Path.GetFileName(path);
                        string errorFullName = errorPath + @"\" + tarFName;
                        if (File.Exists(path))
                        {
                            if (!Directory.Exists(errorPath))
                                Directory.CreateDirectory(errorPath);
                            if (File.Exists(errorFullName))
                                File.Delete(errorFullName);
                            File.Move(path, errorFullName);
                        }
                    }
                    catch (Exception ex)
                    {
                        WriteLog(ex.ToString());
                    }
                    continue;
                }
                if (!Directory.Exists(singlePath))
                    Directory.CreateDirectory(singlePath);

                string pptTitle = "" + fileNameWithoutExtension + " - Microsoft PowerPoint";
                IntPtr m_hGameWnd = User32API.FindWindow("PPTFrameClass", null);
                if (m_hGameWnd != IntPtr.Zero)
                {
                    StringBuilder s = new StringBuilder(512);
                    int intTitle = User32API.GetWindowText(m_hGameWnd, s, s.Capacity); //把this.handle换成你需要的句柄  
                    string strTitle = intTitle.ToString();
                    if (strTitle.Contains("受保护的视图"))
                    {
                        if (File.Exists(path))
                        {
                            if (!Directory.Exists(errorPath))
                                Directory.CreateDirectory(errorPath);
                            string tarFName = Path.GetFileName(path);
                            string errorFullName = errorPath + @"\" + tarFName;
                            File.Move(path, errorFullName);
                        }
                        //MouseClick(105, 80);
                        continue;
                    }
                }
                SetUploadDetails();//生成一个txt,保存上传信息

                TextReplacement();//第一步，文字替换

                ChangeToPic(path);//第二步，转换成图片

                ppt.PPTClose();//关闭PPT，释放资源

                ToCompressPPT(path);//第三步，压缩PPT

                PPTToFlash(path);//第四步，PPT转换成flash

                if (File.Exists(path))
                {
                    if (!Directory.Exists(successPath))
                        Directory.CreateDirectory(successPath);
                    string starFName = Path.GetFileName(path);
                    string successFullName = successPath + @"\" + starFName;
                    if (File.Exists(successFullName))
                        File.Delete(successFullName);
                    File.Move(path, successFullName);
                }
            }
        }
        /// <summary>
        /// 判断是否只读
        /// </summary>
        public bool IsOnlyRead(string pptPath)
        {
            RECT rt = new RECT();

            string pptTitle = "" + fileNameWithExtension + " - Microsoft PowerPoint";
            IntPtr m_hGameWnd = User32API.FindWindow(pptClassName, null);
            if (m_hGameWnd == IntPtr.Zero)
            {
                ppt = new PPTHelper();
                ppt.PPTOpen(pptPath);
                m_hGameWnd = User32API.FindWindow(null, pptTitle);
            }

            pptTitle = "Microsoft Office 激活向导";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd != IntPtr.Zero)
            {
                MessageBox.Show("请将Microsoft Office激活", "提示");
            }
            Thread.Sleep(1000 * 3);
            pptTitle = "" + fileNameWithExtension + " - Microsoft PowerPoint";
            m_hGameWnd = User32API.FindWindow(pptClassName, null);
            User32API.SwitchToThisWindow(m_hGameWnd, true);
            //User32API.GetWindowRect(m_hGameWnd, out rt);
            User32API.MoveWindow(m_hGameWnd, 0, 0, 1300, 800, true);//拖动到左上角
            //MouseCliceBackGround(m_hGameWnd, 570, 50);
            MouseClick(570, 50);
            MouseClick(100, 95);

            pptTitle = "Checking for updates";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd != IntPtr.Zero)
            {
                MessageBox.Show("请将iSpring插件自动更新关闭", "提示");
            }
            Thread.Sleep(1000 * 1);
            pptTitle = "发布为Flash";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd == IntPtr.Zero)
            {
                pptTitle = "iSpring Free";
                m_hGameWnd = User32API.FindWindow(null, pptTitle);
                User32API.SwitchToThisWindow(m_hGameWnd, true);
                User32API.GetWindowRect(m_hGameWnd, out rt);
                User32API.MoveWindow(m_hGameWnd, 0, 0, rt.Width, rt.Height, true);//拖动到左上角
                MouseClick(420, 160);
                ppt.PPTClose();
                return true;
            }
            else
            {
                User32API.SwitchToThisWindow(m_hGameWnd, true);
                User32API.GetWindowRect(m_hGameWnd, out rt);
                User32API.MoveWindow(m_hGameWnd, 0, 0, rt.Width, rt.Height, true);//拖动到左上角
                Thread.Sleep(1000 * 1);
                MouseClick(720, 575);
                return false;
            }
        }
        /// <summary>
        /// 设置上传信息文本
        /// </summary>
        public void SetUploadDetails()
        {
            string tarPath = targetPath + @"\" + fileNameWithoutExtension + @"\";
            #region
            string column = "";//所属栏目
            string style = "1";//所属风格
            string title = fileNameWithoutExtension;//标题
            string keyWords = "";//关键词
            string relatedSearch = "";//相关搜索词
            string smallPicPath = tarPath + fileNameWithoutExtension + "SL.jpg";//缩略图
            string videoPath = tarPath + fileNameWithoutExtension + ".swf";//视频
            string zipPath = tarPath + fileNameWithoutExtension + ".zip";//文件
            string lognPicPath = tarPath + fileNameWithoutExtension + ".jpg";//长图
            string price = "0";//价格
            string soft = "PowerPoint(2010)";//软件
            string page = ppt.PageNum().ToString();//页数
            string proportion = "16:9";//比例
            string score = "5";//星级
            string free_status = "1";//是否免费
            string is_tuijian = "1";//是否推荐
            string notes = "2";//是否包含演讲稿
            string complete = "2";//内容是否完整
            string sort = "1";//排序
            string describe = "";//描述
            string source_describe = "该资源来自用户分享，如果损害了你的权利，请联系网站客服处理。";//内容来源说明

            string pathc = @"C:\123456.txt", paths = @"C:\123456.txt", pathk = @"C:\123456.txt", pathsr = @"C:\123456.txt";

            #endregion

            if (!File.Exists(defaultSettingPath))
            {
                MessageBox.Show("默认配置文件不存在！", "提示");
                return;
            }

            string deStr = File.ReadAllText(defaultSettingPath, Encoding.UTF8);
            string[] deStrArr = SplitByString(deStr, "\r\n");
            for (int i = 0; i < deStrArr.Length; i++)
            {
                if (!string.IsNullOrEmpty(deStrArr[i]))
                {
                    string[] kvArr = SplitByString(deStrArr[i], "->");
                    if (!string.IsNullOrEmpty(kvArr[0]))
                    {
                        switch (kvArr[0])
                        {
                            case "colum":
                                {
                                    pathc = kvArr[1];
                                    break;
                                }
                            case "style":
                                {
                                    paths = kvArr[1];
                                    break;
                                }
                            case "keyword":
                                {
                                    pathk = kvArr[1];
                                    break;
                                }
                            case "search":
                                {
                                    pathsr = kvArr[1];
                                    break;
                                }
                            case "price":
                                {
                                    price = kvArr[1];
                                    break;
                                }
                            case "software":
                                {
                                    soft = kvArr[1];
                                    break;
                                }
                            case "proportion":
                                {
                                    proportion = kvArr[1];
                                    break;
                                }
                            case "score":
                                {
                                    score = kvArr[1];
                                    break;
                                }
                            case "isfree":
                                {
                                    free_status = kvArr[1];
                                    break;
                                }
                            case "istuijian":
                                {
                                    is_tuijian = kvArr[1];
                                    break;
                                }
                            case "notes":
                                {
                                    notes = kvArr[1];
                                    break;
                                }
                            case "complete":
                                {
                                    complete = kvArr[1];
                                    break;
                                }
                            case "sort":
                                {
                                    sort = kvArr[1];
                                    break;
                                }
                            default:
                                break;
                        }
                    }
                }
            }
            keyWordTemp = string.Empty;
            if (!File.Exists(pathc))
            {
                MessageBox.Show("栏目配置文件不存在！", "提示");
                return;
            }
            else
            {
                string tempc = GetResult(pathc, "c");
                if (tempc == "其他" || tempc == "")
                    column = "25";
                else
                    column = tempc;
            }
            if (!File.Exists(paths))
            {
                MessageBox.Show("风格配置文件不存在！", "提示");
                return;
            }
            else
            {
                string temps = GetResult(paths, "s");
                if (temps == "其他" || temps == "")
                    style = "16";
                else
                    style = temps;
            }
            keyWords = keyWordTemp;

            string txtStr = @"menu_id->" + column + "\r\nstyle_ids->" + style + "\r\ntitle->" + title + "\r\nkeyword->" + keyWords
                + "\r\nrelated_search->" + relatedSearch + "\r\nsmallPicPath->" + smallPicPath + "\r\nvideoPath->" + videoPath
                + "\r\nzipPath->" + zipPath + "\r\nlognPicPath->" + lognPicPath + "\r\nprice->" + price + "\r\nsoft->" + soft
                 + "\r\npage->" + page + "\r\nproportion->" + proportion + "\r\nscore->" + score + "\r\nfree_status->" + free_status
                 + "\r\nis_tuijian->" + is_tuijian + "\r\nnotes->" + notes + "\r\ncomplete->" + complete + "\r\nsort->" + sort
                 + "\r\ndescribe->" + describe + "\r\nsource_describe->" + source_describe;
            string tarTxtPath = tarPath + fileNameWithoutExtension + ".txt";
            File.WriteAllText(tarTxtPath, txtStr, Encoding.Default);
        }
        /// <summary> /// 加密字符串   
        /// </summary>  
        /// <param name="str">要加密的字符串</param>  
        /// <returns>加密后的字符串</returns>  
        /// <summary>
        /// 压缩ppt
        /// </summary>
        /// <param name="path"></param>
        public void ToCompressPPT(string path)
        {
            ZipHelper zh = new ZipHelper();
            string tarfile = targetPath + @"\" + fileNameWithoutExtension + @"\" + fileNameWithoutExtension + ".zip";
            zh.ZipFile(path, tarfile);
        }
        /// <summary>
        /// 文字替换
        /// </summary>
        public void TextReplacement()
        {
            if (oldNewList.Count > 0)
            {
                int index = 0;
                Task[] taskArr = new Task[oldNewList.Count];
                foreach (var oldNew in oldNewList)
                {
                    taskArr[index] = Task.Factory.StartNew(() => ppt.ReplaceText(oldNew));
                    index++;
                }
                Task.WaitAll(taskArr);
            }
        }
        /// <summary>
        /// ppt转换成图片
        /// </summary>
        /// <param name="path"></param>
        public void ChangeToPic(string path)
        {
            if (!File.Exists(path))
                return;

            string picOutPath = baseDir + @"TEMPFOLDER\";
            if (!Directory.Exists(picOutPath))
                Directory.CreateDirectory(picOutPath);
            DelectDir(picOutPath);
            ppt.ConvertPics(path, picOutPath, totalWidth, 220, space);
            StitchingPictures(picOutPath);
            PicThumbnail();
        }
        /// <summary>
        /// 图片拼接(大图)
        /// </summary>
        /// <param name="targetPath"></param>
        public void StitchingPictures(string picOutPath, int rowCount = 2)
        {
            GetFiles(picOutPath, ".JPG");
            int bgW = totalWidth - 1;
            int bgH = 0;
            Bitmap picTemp = null;

            picTemp = new Bitmap(picPathList[0]);
            bgH += picTemp.Height;//第一张图的高度

            if ((picPathList.Count - 1) % 2 == 1)
            {
                picTemp = new Bitmap(picPathList[1]);
                bgH += ((picPathList.Count - 1) / 2) * picTemp.Height;
                picTemp = new Bitmap(picPathList[picPathList.Count - 1]);
                bgH += picTemp.Height;

                bgH += space * ((picPathList.Count - 1) / 2 + 1);//间距
            }
            else
            {
                picTemp = new Bitmap(picPathList[1]);
                bgH += ((picPathList.Count - 1) / 2) * picTemp.Height;

                bgH += space * ((picPathList.Count - 1) / 2);//间距
            }


            Bitmap bgImg = new Bitmap(bgW, bgH);
            Graphics g = Graphics.FromImage(bgImg);//最终的背景图
            g.Clear(Color.White);

            Bitmap firstPic = new Bitmap(picPathList[0]);
            g.DrawImage(firstPic, 0, 0, firstPic.Width, firstPic.Height);//拼第一张图

            for (int i = 0; i < (picPathList.Count - 1) / rowCount; i++)//拼接两张小图
            {
                Bitmap bgImgR = new Bitmap(bgW, 220);
                Graphics gr = Graphics.FromImage(bgImgR);
                gr.Clear(Color.White);

                Bitmap picTemp1 = new Bitmap(picPathList[(i * 2 - 1) + 2]);
                Bitmap picTemp2 = new Bitmap(picPathList[(i * 2 - 1) + 3]);
                gr.DrawImage(picTemp1, 0, 0, picTemp1.Width, picTemp1.Height);
                gr.DrawImage(picTemp2, picTemp1.Width + space, 0, picTemp2.Width, picTemp2.Height);
                gr.Dispose();
                g.DrawImage(bgImgR, 0, 440 + 5 + i * (bgImgR.Height + space), bgImgR.Width, bgImgR.Height);
            }
            if ((picPathList.Count - 1) % rowCount == 1)
            {
                Bitmap lastPic = new Bitmap(picPathList[picPathList.Count - 1]);
                g.DrawImage(lastPic, 0, bgH - 440, lastPic.Width, lastPic.Height);
            }

            g.Dispose();

            bgImg.Save(singlePath + fileNameWithoutExtension + ".JPG");
        }
        /// <summary>
        /// 缩略图  299*407
        /// </summary>
        public void PicThumbnail()
        {
            if (picPathList.Count >= 7)
            {
                List<string> thumbnailList = new List<string>();
                for (int i = 0; i < 7; i++)
                {
                    thumbnailList.Add(picPathList[i]);
                }
                int bgtW = 780;
                int bgtH = 7;

                Bitmap picThumbnail = null;
                picThumbnail = new Bitmap(thumbnailList[0]);
                bgtH += picThumbnail.Height;//第一张图的高度
                picThumbnail = new Bitmap(thumbnailList[1]);
                bgtH += picThumbnail.Height * 3;

                Bitmap bgTImg = new Bitmap(bgtW, bgtH);
                Graphics g = Graphics.FromImage(bgTImg);//最终的背景图
                g.Clear(Color.White);

                Bitmap firstTPic = new Bitmap(thumbnailList[0]);
                g.DrawImage(firstTPic, 0, 0, firstTPic.Width, firstTPic.Height);//拼第一张图

                for (int i = 0; i < (thumbnailList.Count - 1) / 2; i++)
                {
                    Bitmap bgImgT = new Bitmap(bgtW, 220);
                    Graphics gr = Graphics.FromImage(bgImgT);
                    gr.Clear(Color.White);

                    Bitmap picTemp1 = new Bitmap(picPathList[(i * 2 - 1) + 2]);
                    Bitmap picTemp2 = new Bitmap(picPathList[(i * 2 - 1) + 3]);
                    gr.DrawImage(picTemp1, 0, 0, picTemp1.Width, picTemp1.Height);
                    gr.DrawImage(picTemp2, picTemp1.Width + space, 0, picTemp2.Width, picTemp2.Height);
                    gr.Dispose();
                    g.DrawImage(bgImgT, 0, 440 + 5 + i * (bgImgT.Height + space), bgImgT.Width, bgImgT.Height);
                }
                g.Dispose();
                Bitmap bitmapSL = null;
                Zoom(bgTImg, scale, scale, out bitmapSL,ZoomType.NearestNeighborInterpolation);
                bitmapSL.Save(singlePath + fileNameWithoutExtension + "SL.JPG");
            }
        }
        public enum ZoomType { NearestNeighborInterpolation, BilinearInterpolation }
        /// <summary>
        /// 图像缩放
        /// </summary>
        /// <param name="srcBmp">原始图像</param>
        /// <param name="width">目标图像宽度</param>
        /// <param name="height">目标图像高度</param>
        /// <param name="dstBmp">目标图像</param>
        /// <param name="GetNearOrBil">缩放选用的算法</param>
        /// <returns>处理成功 true 失败 false</returns>
        public bool Zoom(Bitmap srcBmp, double ratioW, double ratioH, out Bitmap dstBmp, ZoomType zoomType)
        {//ZoomType为自定义的枚举类型
            if (srcBmp == null)
            {
                dstBmp = null;
                return false;
            }
            //若缩放大小与原图一样，则返回原图不做处理
            if ((ratioW == 1.0) && ratioH == 1.0)
            {
                dstBmp = new Bitmap(srcBmp);
                return true;
            }
            //计算缩放高宽
            double height = ratioH * (double)srcBmp.Height;
            double width = ratioW * (double)srcBmp.Width;
            dstBmp = new Bitmap((int)width, (int)height);

            BitmapData srcBmpData = srcBmp.LockBits(new Rectangle(0, 0, srcBmp.Width, srcBmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            BitmapData dstBmpData = dstBmp.LockBits(new Rectangle(0, 0, dstBmp.Width, dstBmp.Height), ImageLockMode.ReadWrite, PixelFormat.Format24bppRgb);
            unsafe
            {
                byte* srcPtr = null;
                byte* dstPtr = null;
                int srcI = 0;
                int srcJ = 0;
                double srcdI = 0;
                double srcdJ = 0;
                double a = 0;
                double b = 0;
                double F1 = 0;//横向插值所得数值
                double F2 = 0;//纵向插值所得数值
                if (zoomType == ZoomType.NearestNeighborInterpolation)
                {//邻近插值法

                    for (int i = 0; i < dstBmp.Height; i++)
                    {
                        srcI = (int)(i / ratioH);//srcI是此时的i对应的原图像的高
                        srcPtr = (byte*)srcBmpData.Scan0 + srcI * srcBmpData.Stride;
                        dstPtr = (byte*)dstBmpData.Scan0 + i * dstBmpData.Stride;
                        for (int j = 0; j < dstBmp.Width; j++)
                        {
                            dstPtr[j * 3] = srcPtr[(int)(j / ratioW) * 3];//j / ratioW求出此时j对应的原图像的宽
                            dstPtr[j * 3 + 1] = srcPtr[(int)(j / ratioW) * 3 + 1];
                            dstPtr[j * 3 + 2] = srcPtr[(int)(j / ratioW) * 3 + 2];
                        }
                    }
                }
                else if (zoomType == ZoomType.BilinearInterpolation)
                {//双线性插值法
                    byte* srcPtrNext = null;
                    for (int i = 0; i < dstBmp.Height; i++)
                    {
                        srcdI = i / ratioH;
                        srcI = (int)srcdI;//当前行对应原始图像的行数
                        srcPtr = (byte*)srcBmpData.Scan0 + srcI * srcBmpData.Stride;//指原始图像的当前行
                        srcPtrNext = (byte*)srcBmpData.Scan0 + (srcI + 1) * srcBmpData.Stride;//指向原始图像的下一行
                        dstPtr = (byte*)dstBmpData.Scan0 + i * dstBmpData.Stride;//指向当前图像的当前行
                        for (int j = 0; j < dstBmp.Width; j++)
                        {
                            srcdJ = j / ratioW;
                            srcJ = (int)srcdJ;//指向原始图像的列
                            if (srcdJ < 1 || srcdJ > srcBmp.Width - 1 || srcdI < 1 || srcdI > srcBmp.Height - 1)
                            {//避免溢出（也可使用循环延拓）
                                dstPtr[j * 3] = 255;
                                dstPtr[j * 3 + 1] = 255;
                                dstPtr[j * 3 + 2] = 255;
                                continue;
                            }
                            a = srcdI - srcI;//计算插入的像素与原始像素距离（决定相邻像素的灰度所占的比例）
                            b = srcdJ - srcJ;
                            for (int k = 0; k < 3; k++)
                            {//插值    公式：f(i+p,j+q)=(1-p)(1-q)f(i,j)+(1-p)qf(i,j+1)+p(1-q)f(i+1,j)+pqf(i+1, j + 1)
                                F1 = (1 - b) * srcPtr[srcJ * 3 + k] + b * srcPtr[(srcJ + 1) * 3 + k];
                                F2 = (1 - b) * srcPtrNext[srcJ * 3 + k] + b * srcPtrNext[(srcJ + 1) * 3 + k];
                                dstPtr[j * 3 + k] = (byte)((1 - a) * F1 + a * F2);
                            }
                        }
                    }
                }
            }
            srcBmp.UnlockBits(srcBmpData);
            dstBmp.UnlockBits(dstBmpData);
            return true;
        }
        /// <summary>
        /// ppt转换成flash
        /// </summary>
        /// <param name="path"></param>
        public void PPTToFlash(string path)
        {
            string tempPPTPath = baseDir + @"PPTTEMP\";
            if (!Directory.Exists(tempPPTPath))
                Directory.CreateDirectory(tempPPTPath);
            DelectDir(tempPPTPath);
            AutoClick(tempPPTPath, path);
        }
        /// <summary>
        /// 图像识别，驱动点击
        /// </summary>
        public void AutoClick(string tempPPTPath, string pptPath)
        {
            RECT rt = new RECT();
            KillProcess("POWERPNT");

            string pptTitle = "" + fileNameWithExtension + " - Microsoft PowerPoint";
            IntPtr m_hGameWnd = User32API.FindWindow(pptClassName, null);
            if (m_hGameWnd == IntPtr.Zero)
            {
                ppt = new PPTHelper();
                ppt.PPTOpen(pptPath);
                m_hGameWnd = User32API.FindWindow(null, pptTitle);
            }

            Thread.Sleep(1000 * 5);

            pptTitle = "Microsoft Office 激活向导";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd != IntPtr.Zero)
            {
                MessageBox.Show("请将Microsoft Office激活", "提示");
            }

            pptTitle = "" + fileNameWithExtension + " - Microsoft PowerPoint";
            m_hGameWnd = User32API.FindWindow(pptClassName, null);
            User32API.SwitchToThisWindow(m_hGameWnd, true);
            //User32API.GetWindowRect(m_hGameWnd, out rt);
            User32API.MoveWindow(m_hGameWnd, 0, 0, 1300, 800, true);//拖动到左上角

            MouseClick(570, 50);
            MouseClick(80, 95);
            //Clipboard.Clear();
            Thread.Sleep(1000);
            pptTitle = "Checking for updates";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd != IntPtr.Zero)
            {
                MessageBox.Show("请将iSpring插件自动更新关闭", "提示");
            }

            pptTitle = "发布为Flash";
            m_hGameWnd = User32API.FindWindow(null, pptTitle);
            if (m_hGameWnd != IntPtr.Zero)
            {
                User32API.SwitchToThisWindow(m_hGameWnd, true);
                User32API.GetWindowRect(m_hGameWnd, out rt);
                User32API.MoveWindow(m_hGameWnd, 0, 0, rt.Width, rt.Height, true);//拖动到左上角
                try
                {
                    Clipboard.Clear();
                    Clipboard.SetText(fileNameWithoutExtension);
                    MouseClick(500, 125);
                    SendKeys.SendWait("^A");
                    Thread.Sleep(500);
                    SendKeys.SendWait("{BACKSPACE}");
                    Thread.Sleep(500);
                    SendKeys.SendWait("^V");  //Ctrl+V 组合键  
                    Thread.Sleep(500);
                    Clipboard.Clear();

                    //User32API.Keybd_event(VirtualKey.BACK, 0, KeyEvent.KEYEVENTF_EXTENDEDKEY, 0);
                    //Thread.Sleep(500);
                    //SendKeys.SendWait(fileNameWithoutExtension);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.ToString());
                }
                try
                {
                    MouseClick(500, 175);
                    Clipboard.Clear();
                    Clipboard.SetText(tempPPTPath);
                    SendKeys.SendWait("^A");
                    Thread.Sleep(500);
                    SendKeys.SendWait("{BACKSPACE}");
                    Thread.Sleep(500);
                    SendKeys.SendWait("^V");  //Ctrl+V 组合键  
                    Thread.Sleep(500);
                    Clipboard.Clear();

                    //MouseClick(500, 175);
                    //User32API.Keybd_event(VirtualKey.BACK, 0, KeyEvent.KEYEVENTF_EXTENDEDKEY, 0);
                    //Thread.Sleep(500);
                    //SendKeys.SendWait(tempPPTPath);
                }
                catch (Exception ex)
                {
                    WriteLog(ex.ToString());
                }
                MouseClick(630, 570);//点击发布

                while (true)
                {
                    Thread.Sleep(1000 * 5);
                    pptTitle = "正在生成Flash影片 {presentation_title}";
                    m_hGameWnd = User32API.FindWindow(null, pptTitle);
                    if (m_hGameWnd == IntPtr.Zero)
                        break;
                }
                KillProcess("360chrome");
                KillProcess("360se");
                KillProcess("explorer");
                KillProcess("chrome");
                KillProcess("iexplore");
            }
            else
            {
                MessageBox.Show("PPT可能无法正常使用iSpring插件", "提示");
                return;
            }

            string tarFlash = targetPath + @"\" + fileNameWithoutExtension + @"\" + fileNameWithoutExtension + ".swf";
            if (File.Exists(tarFlash))
                File.Delete(tarFlash);
            string sPath = tempPPTPath + fileNameWithoutExtension + @"\" + fileNameWithoutExtension + ".swf";
            if (File.Exists(sPath))
            {
                File.Move(sPath, tarFlash);
            }

            if (ppt != null)
                ppt.PPTClose();
            KillProcess("POWERPNT");
        }
        /// <summary>
        /// 鼠标点击
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="index">点击次数</param>
        public void LeftMouseClick(int x, int y, int index = 1)
        {
            User32API.SetCursorPos(x, y);//设置鼠标位置（相对于整个桌面）；
            Thread.Sleep(100);
            for (int i = 0; i < index; i++)
            {
                User32API.MouseEvent(MouseEventType.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
                Thread.Sleep(100);
                User32API.MouseEvent(MouseEventType.MOUSEEVENTF_LEFTUP, x, y, 0, 0);
                Thread.Sleep(100);
            }
        }
        /// <summary>
        /// 鼠标点击
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="option">1：鼠标左键点击，2：鼠标右键点击</param>
        public void MouseClick(int x, int y, int option = 1)
        {
            User32API.SetCursorPos(x, y);//设置鼠标位置（相对于整个桌面）；
            Thread.Sleep(100);
            switch (option)
            {
                case 1:
                    {
                        User32API.MouseEvent(MouseEventType.MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
                        Thread.Sleep(100);
                        User32API.MouseEvent(MouseEventType.MOUSEEVENTF_LEFTUP, x, y, 0, 0);
                        Thread.Sleep(100);
                        break;
                    }
                case 2:
                    {
                        User32API.MouseEvent(MouseEventType.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0);
                        Thread.Sleep(100);
                        User32API.MouseEvent(MouseEventType.MOUSEEVENTF_RIGHTUP, x, y, 0, 0);
                        Thread.Sleep(100);
                        break;
                    }
                default:
                    break;
            }
            Thread.Sleep(1000 * 1);
        }

        public void MouseCliceBackGround(IntPtr hwnd, int x, int y)
        {
            int WM_LBUTTONDOWN = 0x201;//按下
            int WM_LBUTTONUP = 0x202;//弹起
            System.Drawing.Point point = new System.Drawing.Point(x, y);
            User32API.SendMessageA(hwnd, WM_LBUTTONDOWN, point.X, point.Y);
            User32API.SendMessageA(hwnd, WM_LBUTTONUP, point.X, point.Y);
        }
        /// <summary>
        /// 判断窗口是否在运行
        /// </summary>
        /// <param name="WindowsTitle">窗口的标题名称</param>
        /// <returns></returns>
        public bool IsRuning(string WindowsTitle)
        {
            IntPtr WindowHandle = User32API.FindWindow(null, WindowsTitle);//窗口句柄
            if (WindowHandle == IntPtr.Zero)
            {
                return false;
            }
            return true;
        }
        /// <summary>
        /// 杀死进程
        /// </summary>
        /// <param name="pName">进程名</param>
        public void KillProcess(string pName)
        {
            Process[] process;//创建一个PROCESS类数组
            process = Process.GetProcesses();//获取当前任务管理器所有运行中程序
            foreach (Process proces in process)//遍历
            {
                try
                {
                    if (proces.ProcessName == pName)
                    {
                        proces.Kill();
                    }
                }
                catch (Exception ex) { }
            }
        }
        /// <summary>
        /// 将字符串转化成键值对集合
        /// </summary>
        public void StrToKeyValue()
        {
            if (!string.IsNullOrEmpty(wordStr))
            {
                oldNewList.Clear();
                string[] oldNewStrArr = SplitByString(wordStr, "\r\n");
                foreach (var oldNewStr in oldNewStrArr)
                {
                    if (oldNewStr.Contains("-"))
                    {
                        string[] oldNewArr = oldNewStr.Split('-');
                        oldNewList.Add(new KeyValuePair<string, string>(oldNewArr[0], oldNewArr[1]));
                    }
                }
            }
        }
        /// <summary>
        /// 读取目录下的文件
        /// </summary>
        /// <param name="extNameStr">扩展名</param>
        public void GetFiles(string fileDir, string extNameStr)
        {
            if (!string.IsNullOrEmpty(fileDir))
            {
                if (!Directory.Exists(fileDir))
                {
                    WriteLog("文件夹不存在！");
                    return;
                }
                if (extNameStr == ".pptx" || extNameStr == ".ppt")
                    pptPathList.Clear();
                if (extNameStr.ToUpper() == ".JPG")
                    picPathList.Clear();
                DirectoryInfo mydir = new DirectoryInfo(fileDir);
                int index = 0;
                foreach (FileSystemInfo fsi in mydir.GetFileSystemInfos())
                {
                    if (fsi is FileInfo)
                    {
                        string extName = Path.GetExtension(fsi.FullName); //获取扩展名  
                        if (extName.ToUpper() == extNameStr.ToUpper() || extName.ToLower() == ".ppt")
                        {
                            if (extNameStr == ".pptx" || extNameStr == ".ppt")
                            {
                                pptPathList.Add(fsi.FullName);
                            }
                            else if (extNameStr.ToUpper() == ".JPG")
                            {
                                index++;
                                picPathList.Add(fileDir + index + ".JPG");
                            }
                        }
                    }
                }
            }
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
        /// 删除指定目录下的所有文件和子目录
        /// </summary>
        /// <param name="srcPath"></param>
        public static void DelectDir(string srcPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(srcPath);
                FileSystemInfo[] fileinfo = dir.GetFileSystemInfos();  //返回目录中所有文件和子目录
                foreach (FileSystemInfo i in fileinfo)
                {
                    if (i is DirectoryInfo)            //判断是否文件夹
                    {
                        DirectoryInfo subdir = new DirectoryInfo(i.FullName);
                        subdir.Delete(true);          //删除子目录和文件
                    }
                    else
                    {
                        File.Delete(i.FullName);      //删除指定文件
                    }
                }
            }
            catch (Exception e)
            {
            }
        }
        /// <summary>
        /// 日志打印
        /// </summary>
        /// <param name="log"></param>
        public static void WriteLog(string log)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "log\\";//日志文件夹
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)//判断文件夹是否存在
                dir.Create();//不存在则创建

            FileInfo[] subFiles = dir.GetFiles();//获取该文件夹下的所有文件
            foreach (FileInfo f in subFiles)
            {
                string fname = Path.GetFileNameWithoutExtension(f.FullName); //获取文件名，没有后缀
                DateTime start = Convert.ToDateTime(fname);//文件名转换成时间
                DateTime end = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));//获取当前日期
                TimeSpan sp = end.Subtract(start);//计算时间差
                if (sp.Days > 30)//大于30天删除
                    f.Delete();
            }

            string logName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";//日志文件名称，按照当天的日期命名
            string fullPath = path + logName;//日志文件的完整路径
            string contents = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " -> " + log + "\r\n";//日志内容

            File.AppendAllText(fullPath, contents, Encoding.UTF8);//追加日志
        }

        public string GetResult(string path, string option)
        {
            string res = string.Empty;
            if (!File.Exists(path))
            {
                MessageBox.Show("默认配置文件不存在！", "提示");
                return null;
            }

            string deStr = File.ReadAllText(path, Encoding.Default);
            string[] deStrArr = SplitByString(deStr, "\r\n");
            for (int i = 0; i < deStrArr.Length; i++)
            {
                if (!string.IsNullOrEmpty(deStrArr[i]))
                {
                    string[] dkv = SplitByString(deStrArr[i], "=");
                    if (DelWords(dkv[1]))
                    {
                        if (option == "c")
                            res = MatchId(dkv[0], option);
                        if (option == "s")
                            res += MatchId(dkv[0], option) + ",";
                        keyWordTemp += dkv[1].Replace("\r\n", "、");
                    }
                }
            }
            return res;
        }
        /// <summary>
        /// 判断标题是否包含关键词
        /// </summary>
        /// <param name="wordStr"></param>
        /// <returns></returns>
        public bool DelWords(string wordStr)
        {
            string[] wordArr = wordStr.Split('、');
            for (int i = 0; i < wordArr.Length; i++)
            {
                if (!string.IsNullOrEmpty(wordArr[i]))
                {
                    if (fileNameWithoutExtension.Contains(wordArr[i]))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        /// <summary>
        /// 匹配ID
        /// </summary>
        /// <param name="key"></param>
        /// <param name="option"></param>
        /// <returns></returns>
        public string MatchId(string key, string option)
        {
            Dictionary<string, string> cloumKV = new Dictionary<string, string>();
            #region
            cloumKV.Add("商业计划书", "15");
            cloumKV.Add("教育培训", "16");
            cloumKV.Add("公司简介", "17");
            cloumKV.Add("政府党建", "18");
            cloumKV.Add("产品发布", "19");
            cloumKV.Add("节假节日", "20");
            cloumKV.Add("年会颁奖", "21");
            cloumKV.Add("简历竞聘", "22");
            cloumKV.Add("医学护理", "23");
            cloumKV.Add("婚庆爱情", "24");
            cloumKV.Add("其他", "25");
            cloumKV.Add("汇报总结", "7");
            cloumKV.Add("论文答辩", "8");
            #endregion
            Dictionary<string, string> styleKV = new Dictionary<string, string>();
            #region
            styleKV.Add("简约", "1");
            styleKV.Add("商务", "2");
            styleKV.Add("中国风", "3");
            styleKV.Add("小清新", "4");
            styleKV.Add("杂志风", "15");
            styleKV.Add("欧美风", "8");
            styleKV.Add("微粒体", "9");
            styleKV.Add("手绘卡通", "10");
            styleKV.Add("黑板风", "14");
            styleKV.Add("扁平化", "12");
            styleKV.Add("炫酷", "13");
            styleKV.Add("其他", "16");
            #endregion
            try
            {
                switch (option)
                {
                    case "c":
                        {
                            if (string.IsNullOrEmpty(cloumKV[key]))
                                return "25";
                            return cloumKV[key];
                        }
                    case "s":
                        {
                            if (string.IsNullOrEmpty(styleKV[key]))
                                return "16";
                            return styleKV[key];
                        }
                    default:
                        return null;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("栏目或风格是否已经变动，如有变动请与开发者联系。\r\nQQ：1610779207", "提示");
                return null;
            }
        }
    }
}
