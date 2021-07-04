using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using Microsoft.Office.Interop.Word;
using System.Threading;
using WindowsFormsApp1;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        public Word.Application m_app;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            m_app = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Selection select = m_app.Selection;
            MessageBox.Show(select.Text);
            MessageBox.Show("你好");
            Word.Document doc = m_app.ActiveDocument;
            int nCount = m_app.ActiveDocument.Styles.Count;
            for (int j = 1; j <= nCount; j++)
            {
                Word.Style stl = m_app.ActiveDocument.Styles[j];
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            Word.Paragraphs paras = m_app.ActiveDocument.Paragraphs;
            int count = paras.Count;
            Word.Paragraph para;
            try
            {
                var dg = m_app.Dialogs[WdWordDialog.wdDialogFormatFont];
                dg.Show();
                var c = dg.Application.Selection.Range.Font.Color;
                //m_app.ActiveDocument.Research.Query("tets", "犯得");
                //return;
                // var selList = m_app.ActiveDocument.Range().FindAll(edit_1.Text);
                //m_app.Selection.WholeStory Find.Execute(FindText: findText, MatchCase: true);
                //var r = m_app.Selection.Range.FindAll("犯得").Select(s => new { s.Start, s.End }).ToList();
                //colorDialog1.ShowDialog();
                System.Threading.Thread.Sleep(100);
                //foreach (var item in selList)
                //{
                //    item.Select();

                //    //break;
                //    // m_app.Selection.SetRange(item.Start, item.End);
                //    Word.Range rg = m_app.Selection.Range;
                //    rg.Font.Color = c;
                //}


                //Form1 form1 = new Form1();
                //form1.Show();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

            return;
            for (int i = 1; i <= count; i++)
            {
                MessageBox.Show(paras[i].Range.Text);
                switch (paras[i].OutlineLevel)
                {
                    // 一级标题
                    case Word.WdOutlineLevel.wdOutlineLevel1:
                        {
                            para = paras[i];
                            para.Range.Font.Bold = 0;// 不加粗
                            para.Range.Font.Size = float.Parse("15");// 小三
                            para.Range.Font.Name = "黑体";// 黑体
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;// 黑色
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;// 居中
                            para.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;// 固定值
                            para.Range.ParagraphFormat.LineSpacing = float.Parse("20");// 20 磅
                        }
                        break;

                    // 二级标题
                    case Word.WdOutlineLevel.wdOutlineLevel2:
                        {
                            para = paras[i];
                            para.Range.Font.Bold = 0;// 不加粗
                            para.Range.Font.Size = float.Parse("14");// 四号
                            para.Range.Font.Name = "黑体";// 黑体
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;// 黑色
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;// 左对齐
                            para.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;// 固定值
                            para.Range.ParagraphFormat.LineSpacing = float.Parse("20");// 20 磅
                        }
                        break;

                    // 三级标题
                    case Word.WdOutlineLevel.wdOutlineLevel3:
                        {
                            para = paras[i];
                            para.Range.Font.Bold = 0;// 不加粗
                            para.Range.Font.Size = float.Parse("12");// 小四
                            para.Range.Font.Name = "宋体";// 宋体
                            para.Range.Font.Color = Word.WdColor.wdColorBlack;// 黑色
                            para.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;// 左对齐
                            para.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;// 固定值
                            para.Range.ParagraphFormat.LineSpacing = float.Parse("20");// 20 磅
                        }
                        break;
                    default:
                        break;
                }
            }

            //Word.Paragraphs paras = m_app.ActiveDocument.Paragraphs;
            //paras.
            //m_app.Selection.Select();
            //Word.Selection select = m_app.Selection;
            //MessageBox.Show(select.Text);
            //Word.Range rg = select.Range;
            //rg.FindAll("发");

            //rg.Font.Bold = 0; // 不加粗
            //rg.Font.Size = float.Parse("12"); // 小四
            //rg.Font.Name = "宋体";// 宋体
            //rg.Font.Color = Word.WdColor.wdColorBlue;// 字体颜色黑色

            //select.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;// 两端对齐
            //if (select.ParagraphFormat.CharacterUnitFirstLineIndent == float.Parse("2") ||
            //    select.ParagraphFormat.FirstLineIndent == rg.Font.Size * 2)// 缩进了两个字符大小
            //{
            //    return;
            //}

            //select.ParagraphFormat.FirstLineIndent = float.Parse("0");
            //select.ParagraphFormat.IndentFirstLineCharWidth(2);// 首行缩进两个字符  
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            var allRange = m_app.ActiveDocument.Range();
            var selList1 = allRange.SearchRangeInPattern("[\"'](.*?)[\"']");
            var selList2 = allRange.SearchRangeInPattern("[\"“](.*?)[\"”]");
            var selList = selList1.Concat(selList2);
            var dg = m_app.Dialogs[WdWordDialog.wdDialogFormatFont];
            dg.Show();
            var selColor = dg.Application.Selection.Range.Font.Color;
            System.Threading.Thread.Sleep(1000);
            foreach (var item in selList)
            {
                item.Select();

                //break;
                // m_app.Selection.SetRange(item.Start, item.End);
                Word.Range rg = m_app.Selection.Range;
                rg.Font.Color = selColor;
            }

        }

        private void setCheck(string text)
        {
            Word.Paragraphs paras = m_app.ActiveDocument.Paragraphs;
            int count = paras.Count;
            Word.Paragraph para;
            try
            {
                var dg = m_app.Dialogs[WdWordDialog.wdDialogFormatFont];
                dg.Show();
                var c = dg.Application.Selection.Range.Font.Color;
                //m_app.ActiveDocument.Research.Query("tets", "犯得");
                //return;

                var selList = m_app.ActiveDocument.Range().FindAll(text);
                //m_app.Selection.WholeStory Find.Execute(FindText: findText, MatchCase: true);
                //var r = m_app.Selection.Range.FindAll("犯得").Select(s => new { s.Start, s.End }).ToList();
                //colorDialog1.ShowDialog();
                System.Threading.Thread.Sleep(1000);
                foreach (var item in selList)
                {
                    item.Select();

                    //break;
                    // m_app.Selection.SetRange(item.Start, item.End);
                    Word.Range rg = m_app.Selection.Range;
                    rg.Font.Color = c;
                }

                //Form1 form1 = new Form1();
                //form1.Show();

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }

        }


        private void button1_Click_2(object sender, RibbonControlEventArgs e)
        {
            try
            {
                //在打开的窗口中查找about窗口
                Form test = System.Windows.Forms.Application.OpenForms["Form1"];

                //判断对应窗口是否被打开
                if ((test == null) || (test.IsDisposed))
                {
                    Form1 form1 = new Form1();
                    form1.TopMost = true;
                    form1.Location = new System.Drawing.Point(System.Windows.Forms.Screen.PrimaryScreen.WorkingArea.Width - 10 - form1.Width, 10);
                    form1.Show();
                }
                else
                {
                    //如果已经打开了

                    //让其获得焦点
                    test.Activate();
                    //窗口恢复正常
                    test.WindowState = FormWindowState.Normal;
                }

            
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
       
        }
    }
    public enum WdStyleType
    {
        // Paragraph style.
        wdStyleTypeParagraph = 1,

        // Body character style.
        wdStyleTypeCharacter = 2,

        // Table style.
        wdStyleTypeTable = 3,

        // List style.
        wdStyleTypeList = 4,

        // Reserved for internal use.
        wdStyleTypeParagraphOnly = 5,

        // Reserved for internal use.
        wdStyleTypeLinked = 6
    }
}
