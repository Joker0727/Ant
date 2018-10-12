using System;
using System.Collections.Generic;
using System.Diagnostics;
using OFFICECORE = Microsoft.Office.Core;
using POWERPOINT = Microsoft.Office.Interop.PowerPoint;

namespace MyPPT
{
    public class PPTHelper
    {
        #region=========基本的参数信息=======
        public POWERPOINT.Application objApp = null;
        public POWERPOINT.Presentation objPresSet = null;
        public POWERPOINT.SlideShowWindows objSSWs;
        public POWERPOINT.SlideShowTransition objSST;
        public POWERPOINT.SlideShowSettings objSSS;
        public POWERPOINT.SlideRange objSldRng;
        public bool bAssistantOn;

        #endregion
        #region===========操作方法==============
        /// <summary>
        /// 打开PPT文档并播放显示。
        /// </summary>
        /// <param name="filePath">PPT文件路径</param>
        public void PPTOpen(string filePath)
        {
            KillProcess("POWERPNT");
            //防止连续打开多个PPT程序.
            if (this.objApp != null)
            {
                return;
            }
            for (int i = 0; i < 3; i++)
            {
                try
                {
                    objApp = new POWERPOINT.Application();
                    objApp.Visible = OFFICECORE.MsoTriState.msoTrue;
                    //以非只读方式打开,方便操作结束后保存.
                    objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoFalse);
                }
                catch (Exception ex)
                {
                    //MessageBox.Show("错误:" + ex.Message.ToString());
                    try
                    {
                        this.objApp.Quit();
                    }
                    catch (Exception e) { }
                    finally
                    {
                        this.objApp = null;
                        this.objPresSet = null;
                    }
                    KillProcess("POWERPNT");
                    continue;
                }
                break;
            }
        }
        /// <summary>
        /// 自动播放PPT文档.
        /// </summary>
        /// <param name="filePath">PPTy文件路径.</param>
        /// <param name="playTime">翻页的时间间隔.【以秒为单位】</param>
        public void PPTAuto(string filePath, int playTime)
        {
            //防止连续打开多个PPT程序.
            if (this.objApp != null) { return; }
            objApp = new POWERPOINT.Application();
            objPresSet = objApp.Presentations.Open(filePath, OFFICECORE.MsoTriState.msoCTrue, OFFICECORE.MsoTriState.msoFalse, OFFICECORE.MsoTriState.msoFalse);
            // 自动播放的代码（开始）
            int Slides = objPresSet.Slides.Count;
            int[] SlideIdx = new int[Slides];
            for (int i = 0; i < Slides; i++) { SlideIdx[i] = i + 1; };
            objSldRng = objPresSet.Slides.Range(SlideIdx);
            objSST = objSldRng.SlideShowTransition;
            //设置翻页的时间.
            objSST.AdvanceOnTime = OFFICECORE.MsoTriState.msoCTrue;
            objSST.AdvanceTime = playTime;
            //翻页时的特效!
            objSST.EntryEffect = POWERPOINT.PpEntryEffect.ppEffectCircleOut;
            //Prevent Office Assistant from displaying alert messages:
            bAssistantOn = objApp.Assistant.On;
            objApp.Assistant.On = false;
            //Run the Slide show from slides 1 thru 3.
            objSSS = objPresSet.SlideShowSettings;
            objSSS.StartingSlide = 1;
            objSSS.EndingSlide = Slides;
            objSSS.Run();
            //Wait for the slide show to end.
            objSSWs = objApp.SlideShowWindows;
            while (objSSWs.Count >= 1) System.Threading.Thread.Sleep(playTime * 100);
            this.objPresSet.Close();
            this.objApp.Quit();
        }
        /// <summary>
        /// PPT下一页。
        /// </summary>
        public void NextSlide()
        {
            if (this.objApp != null)
                try
                {
                    this.objPresSet.SlideShowWindow.View.Next();
                }
                catch
                { }
        }
        /// <summary>
        /// PPT上一页。
        /// </summary>
        public void PreviousSlide()
        {
            if (this.objApp != null)
                this.objPresSet.SlideShowWindow.View.Previous();
        }
        /// <summary>
        /// 获取幻灯片页数
        /// </summary>
        /// <returns></returns>
        public int PageNum()
        {
            return objPresSet.Slides.Count;
        }
        /// <summary>
        /// 设置边框
        /// </summary>
        public void SetLine()
        {
            int num = PageNum();
            for (int i = 0; i < num; i++)
            {
                if (i > 2)
                {
                    objSldRng = objPresSet.Slides.Range(i);
                    objSldRng.Select();
                    try
                    {
                        objSldRng.Application.ActiveWindow.Selection.SlideRange.Shapes.SelectAll();
                        objSldRng.Application.ActiveWindow.Selection.ShapeRange.Line.Visible = OFFICECORE.MsoTriState.msoFalse;
                    }
                    catch
                    { }
                    //MessageBox.Show("" + i.ToString());


                    //NextSlide();
                }

            }

        }
        /// <summary>
        /// 文字替换
        /// </summary>
        public void ReplaceText(KeyValuePair<string, string> kv)
        {
            string OldText = string.Empty;
            string NewText = string.Empty;

            OldText = kv.Key;
            NewText = kv.Value;

            int num = PageNum();
            for (int j = 1; j <= num; j++)
            {
                POWERPOINT.Slide slide = objPresSet.Slides[j];
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    try
                    {

                        POWERPOINT.Shape shape = slide.Shapes[i];
                        if (shape.TextFrame != null)
                        {
                            POWERPOINT.TextFrame textFrame = shape.TextFrame;
                            string oldText = textFrame.TextRange.Text;
                            string newText = textFrame.TextRange.Text.Replace(OldText, NewText);
                            if (textFrame.TextRange != null && !string.IsNullOrEmpty(oldText))
                            {
                                textFrame.TextRange.Replace(oldText, newText);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
        }
        /// <summary>
        /// 把每张幻灯片转换成图片
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <param name="targetFileType"></param>
        public void ConvertPics(string sourcePath, string picOutPath, int totalWidth, int height, int space, POWERPOINT.PpSaveAsFileType targetFileType = POWERPOINT.PpSaveAsFileType.ppSaveAsJPG, string type = "JPG")
        {
            int width = (totalWidth - space) / 2;
            //   persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue);//整个ppt的文件转换为其他的格式
            objPresSet.Slides[1].Export(picOutPath + "1." + type, type, totalWidth, 440); //将ppt中的某张转换为图片文件
            for (int i = 2; i < objPresSet.Slides.Count + 1; i++)
            {
                try
                {
                    if (!Convert.ToBoolean((objPresSet.Slides.Count - 1) % 2))
                        objPresSet.Slides[i].Export(picOutPath + i + "." + type, type, width, height); //将ppt中的某张转换为图片文件
                    else
                    {
                        if (i == objPresSet.Slides.Count)
                            objPresSet.Slides[i].Export(picOutPath + i + "." + type, type, totalWidth, 440);
                        else
                            objPresSet.Slides[i].Export(picOutPath + i + "." + type, type, width, height);
                    }
                }
                catch (Exception ex) { }
            }

            //persentation.SaveAs(targetPath, targetFileType, Microsoft.Office.Core.MsoTriState.msoTrue); //整个ppt的文件转换为其他的格式           
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
        /// 关闭PPT文档。
        /// </summary>
        public void PPTClose()
        {
            //装备PPT程序。
            if (this.objPresSet != null)
            {
                try
                {
                    this.objPresSet.Save();
                }
                catch (Exception e) { }
                finally
                {
                    objPresSet.Close();
                    objPresSet = null;
                }
            }
            if (this.objApp != null)
            {
                try
                {
                    this.objApp.Quit();
                }
                catch (Exception ex) { }
                finally
                {
                    objApp = null;
                }
            }
            KillProcess("POWERPNT");
            GC.Collect();
        }
        #endregion
    }
}
