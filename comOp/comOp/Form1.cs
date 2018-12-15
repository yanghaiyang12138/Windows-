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
using MsWord = Microsoft.Office.Interop.Word;
using MsExcel = Microsoft.Office.Interop.Excel;

namespace comOp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.ShowDialog();
            string txt_file = opf.FileName;

            StreamReader sr = new StreamReader(txt_file);
            //List<string> sl = new List<string>();
            //string line;
            //while ((line = sr.ReadLine()) != null)
            //{
            //    sl.Add(line);
            //}


            MsWord.Application oWordApplic;
            MsWord.Document oDoc;
            try
            {
                string doc_file_name = @"E:\QQFileRec\Com实验用到的文件\new.doc";
                if (File.Exists(doc_file_name))
                {
                    File.Delete(doc_file_name);
                }
                oWordApplic = new MsWord.Application();
                object missing = System.Reflection.Missing.Value;

                MsWord.Range curRange;
                object curTxt;
                int curSectionNum = 1;
                oDoc = oWordApplic.Documents.Add(ref missing, ref missing, ref missing, ref missing);
                oDoc.Activate();
                //Console.WriteLine("正在生成文档小节");
                //object section_nextPage = MsWord.WdBreakType.wdSectionBreakNextPage;
                //for (int si = 0; si < 4; si++)
                //{
                //    oDoc.Paragraphs[1].Range.InsertParagraphAfter();
                //    oDoc.Paragraphs[1].Range.InsertBreak(ref section_nextPage);
                //}

                Console.WriteLine("正在插入摘要部分");
                curSectionNum = 1;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curRange.Select();
                string one_str, key_word;
                oWordApplic.Options.Overtype = false;
                MsWord.Selection currentSelection = oWordApplic.Selection;
                if (currentSelection.Type == MsWord.WdSelectionType.wdSelectionNormal)
                {
                    one_str = sr.ReadLine();//标题
                    Console.WriteLine("标题：" + one_str);
                    currentSelection.TypeText(one_str);
                    currentSelection.TypeParagraph();//第一段
                    
                    key_word = sr.ReadLine();//读入关键字
                    Console.WriteLine("关键字：" + key_word);
                   

                    currentSelection.TypeText("摘要：");//添加摘要二字
                    //currentSelection.TypeParagraph();
                    one_str = sr.ReadLine();//读入段落文本
                    while (one_str != null)
                    {
                        currentSelection.TypeText(one_str);
                        currentSelection.TypeParagraph();//添加段落标记
                        one_str = sr.ReadLine();
                    }

                    currentSelection.TypeText("关键字："+key_word);
                    currentSelection.TypeParagraph();//添加段落标记

                }
                sr.Close();

                //标题
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[1].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                Console.WriteLine(curTxt);
                curRange.Font.Name = " 宋体";
                curRange.Font.Size = 22;
                curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;



                //摘要正文
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[2].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                curRange.Select();
                curRange.Font.Name = "宋体";
                curRange.Font.Size = 12;
                oDoc.Sections[curSectionNum].Range.Paragraphs[2].LineSpacingRule = MsWord.WdLineSpacing.wdLineSpaceMultiple;
                oDoc.Sections[curSectionNum].Range.Paragraphs[2].LineSpacing = 15f;
                //oDoc.Sections[curSectionNum].Range.Paragraphs[2].IndentFirstLineCharWidth(2);
                //设置摘要两个字为黑体
                curRange = curRange.Paragraphs[curRange.Paragraphs.Count].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                object range_start, range_end;
                range_start = curRange.Start;
                range_end = curRange.Start + 3;
                curRange = oDoc.Range(ref range_start, ref range_end);
                curTxt = curRange.Text;
                curRange.Select();
                curRange.Font.Bold = 1;
                for (int i = 3; i < oDoc.Sections[curSectionNum].Range.Paragraphs.Count-1; i++)
                {
                    curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[i].Range;
                    curTxt = curRange.Paragraphs[1].Range.Text;
                    curRange.Select();
                    curRange.Font.Name = "宋体";
                    curRange.Font.Size = 12;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacingRule = MsWord.WdLineSpacing.wdLineSpaceMultiple;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacing = 15f;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].IndentFirstLineCharWidth(2);
                }

                //关键字
                oDoc.Sections[curSectionNum].Range.Paragraphs[6].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphLeft;
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[6].Range;
                curRange.Select();
                curTxt = curRange.Paragraphs[1].Range.Text;
                Console.WriteLine(curTxt);
                //curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphLeft;
                curRange.Font.Name = "黑体";
                curRange.Font.Size = 10;
                //设置关键字为黑体
                curRange = curRange.Paragraphs[curRange.Paragraphs.Count].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                //object range_start, range_end;
                range_start = curRange.Start;
                range_end = curRange.Start + 4;
                curRange = oDoc.Range(ref range_start, ref range_end);
                curTxt = curRange.Text;
                curRange.Select();
                curRange.Font.Bold = 1;

                //开始读入正文
                sr = new StreamReader(@"E:\QQFileRec\Com实验用到的文件\content.txt");
                string title,para;
                title = sr.ReadLine();
                currentSelection.TypeText(title);//大标题
                currentSelection.TypeParagraph();

                title = sr.ReadLine();
                currentSelection.TypeText(title);//小标题
                currentSelection.TypeParagraph();

                while ((para = sr.ReadLine()) != null)//正文部分
                {
                    currentSelection.TypeText(para);
                    currentSelection.TypeParagraph();
                }

                sr.Close();

                //设置大标题的格式
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[6].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                Console.WriteLine(curTxt);
                curRange.Font.Name = " 宋体";
                curRange.Font.Size = 18;
                curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;

                //设置小标题的格式
                curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[7].Range;
                curTxt = curRange.Paragraphs[1].Range.Text;
                Console.WriteLine(curTxt);
                curRange.Font.Name = " 宋体";
                curRange.Font.Size = 16;
                curRange.Paragraphs[1].Alignment = MsWord.WdParagraphAlignment.wdAlignParagraphCenter;

                //设置段落
                for (int i = 8; i < oDoc.Sections[curSectionNum].Range.Paragraphs.Count; i++)
                {
                    curRange = oDoc.Sections[curSectionNum].Range.Paragraphs[i].Range;
                    curTxt = curRange.Paragraphs[1].Range.Text;
                    curRange.Select();
                    curRange.Font.Name = "宋体";
                    curRange.Font.Size = 12;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacingRule = MsWord.WdLineSpacing.wdLineSpaceMultiple;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].LineSpacing = 15f;
                    oDoc.Sections[curSectionNum].Range.Paragraphs[i].IndentFirstLineCharWidth(2);
                }


                //文档保存
                //Console.WriteLine("正在更新目录");
                //oDoc.Fields[1].Update();
                Console.WriteLine("正在保存word文档");
                object fileName;
                fileName = doc_file_name;
                oDoc.SaveAs2(ref fileName);
                oDoc.Close();
                Console.WriteLine("正在释放com资源");
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oDoc);
                oDoc = null;
                oWordApplic.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oWordApplic);
                oWordApplic = null;
            }
            catch (Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
            finally
            {
                Console.WriteLine("正在结束word进程");
                System.Diagnostics.Process[] ALLProcess = System.Diagnostics.Process.GetProcesses();
                for(int j = 0; j < ALLProcess.Length; j++)
                {
                    string theProcName = ALLProcess[j].ProcessName;
                    if (string.Compare(theProcName, "WINWORD") == 0)
                    {
                        if (ALLProcess[j].Responding && !ALLProcess[j].HasExited)
                        {
                            ALLProcess[j].Kill();
                        }
                    }
                }
                Console.WriteLine("进程结束");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MsExcel.Application oExcApp;
            MsExcel.Workbook oExcBook;
            try
            {
                //oExcApp = new MsExcel.Application();
                //oExcBook == oExcApp.Workbooks.Add(true);
                //MsExcel
            }catch(Exception e1)
            {
                MessageBox.Show(e1.Message);
            }
        }
    }
}
