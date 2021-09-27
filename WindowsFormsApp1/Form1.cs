using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Spire.Doc;
using Spire.Doc.Documents;
using DocumentFormat.OpenXml.Packaging;
using WindowsFormsApp1.Service;
using Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        Stream stream = new System.IO.MemoryStream();
        public Form1()
        {
            InitializeComponent();
        }


        /// <summary>
        /// 寫入目錄
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>

        private void button1_Click(object sender, EventArgs e)
        {
            using (DocX doc = DocX.Create(stream, DocumentTypes.Document))
            {

                doc.DifferentFirstPage = false;
                var tocSwitches = new Dictionary<TableOfContentsSwitches, string>()
                                {
                                  { TableOfContentsSwitches.O, "1-3"},
                                  { TableOfContentsSwitches.U, ""},
                                  { TableOfContentsSwitches.Z, ""},
                                  { TableOfContentsSwitches.H, ""},
                                };


                doc.InsertTableOfContents("目錄", tocSwitches);

                doc.InsertSectionPageBreak();                
                doc.InsertParagraph("壹、測試內容C#輸出").FontSize(24).Bold(true).Heading(HeadingType.Heading1).Font("標楷體").Color(Color.Black);
                List<testdata> datas = new List<testdata>();
                datas.Add(new testdata { Step = "01", StepDesc = "机台班组人员在SFC上选择要执行的工单，点击工单准备；", InvolvedPost = "机台人员", Remark = "" });
                datas.Add(new testdata { Step = "02", StepDesc = "WIP分析物料位置，获取符合条件的物料所在位置", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "03", StepDesc = "WIP向TMS发送移库任务", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "04", StepDesc = "TMS分析路径；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "05", StepDesc = "TMS向配送员PDA推送任务；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "06", StepDesc = "配送员扫托盘号取货。", InvolvedPost = "配送员", Remark = "" });

                doc.InsertTable(AddTable(datas));

                doc.InsertSectionPageBreak();

                DocX docToMerge = DocX.Load("D:\\source.docx");
                doc.InsertDocument(docToMerge);

                doc.InsertSectionPageBreak();
                doc.InsertParagraph("伍、測試再內容C#輸出").FontSize(24).Bold(true).Heading(HeadingType.Heading1).Font("標楷體").Color(Color.Black);

                doc.InsertTable(AddTable(datas));
                doc.SaveAs("D:\\Test.docx");
            }

            
           




            //Spire.Doc.Document document = new Spire.Doc.Document();
            //document.LoadFromStream(stream,FileFormat.Docx);
            ////設圖片水印
            //PictureWatermark picture = new PictureWatermark();
            //picture.Picture = System.Drawing.Image.FromFile("D:\\252191.jpg");
            //picture.Scaling = 80;
            //document.Watermark = picture;
            //document.SaveToFile("D:\\水印.docx");
        }

       

        public void insert()
        {
            using (DocX doc2 = DocX.Load("D:\\Test.docx"))
            {    
                List<testdata> datas = new List<testdata>();
                datas.Add(new testdata { Step = "01", StepDesc = "机台班组人员在SFC上选择要执行的工单，点击工单准备；", InvolvedPost = "机台人员", Remark = "" });
                datas.Add(new testdata { Step = "02", StepDesc = "WIP分析物料位置，获取符合条件的物料所在位置", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "03", StepDesc = "WIP向TMS发送移库任务", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "04", StepDesc = "TMS分析路径；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "05", StepDesc = "TMS向配送员PDA推送任务；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "06", StepDesc = "配送员扫托盘号取货。", InvolvedPost = "配送员", Remark = "" });

                doc2.InsertTable(AddTable(datas));

                doc2.InsertSectionPageBreak();

                DocX docToMerge = DocX.Load("D:\\source.docx");
                doc2.InsertDocument(docToMerge);

                doc2.InsertSectionPageBreak();
                doc2.InsertParagraph("伍、測試再內容C#輸出").FontSize(24).Bold(true).Heading(HeadingType.Heading1).Font("標楷體").Color(Color.Black);

                doc2.InsertTable(AddTable(datas));
                doc2.SaveAs("D:\\Temp.docx");
            }
        }

        public Xceed.Document.NET.Table AddTable<T>(List<T> obj) where T:class
        {
            DocX doc = DocX.Create("",DocumentTypes.Template); //僅暫存不儲存檔案
            PropertyInfo[] prop = typeof(T).GetProperties();
            Xceed.Document.NET.Table tb = doc.AddTable(obj.Count()+1, prop.Count());
            tb.Design = TableDesign.TableGrid;
            tb.Alignment = Alignment.center;
            foreach (var item in prop.Select((index,value)=>new { index,value}))  //寫入標題
            {
                tb.Rows[0].Cells[item.value].Paragraphs[0].Append(item.index.Name);
            }

            foreach (var row in obj.Select((index,value)=>new { index, value}))
            {
                foreach (var col in prop.Select((index, value) => new { index, value }))  
                {                    
                    tb.Rows[row.value+1].Cells[col.value].Paragraphs[0].Append(col.index.GetValue(row.index).ToString());
                }
            }
            return tb;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (DocX doc = DocX.Create("D:\\TEMP.docx"))
            {
                List<testdata> datas = new List<testdata>();
                datas.Add(new testdata { Step = "01", StepDesc = "机台班组人员在SFC上选择要执行的工单，点击工单准备；", InvolvedPost = "机台人员", Remark = "" });
                datas.Add(new testdata { Step = "02", StepDesc = "WIP分析物料位置，获取符合条件的物料所在位置", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "03", StepDesc = "WIP向TMS发送移库任务", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "04", StepDesc = "TMS分析路径；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "05", StepDesc = "TMS向配送员PDA推送任务；", InvolvedPost = "", Remark = "" });
                datas.Add(new testdata { Step = "06", StepDesc = "配送员扫托盘号取货。", InvolvedPost = "配送员", Remark = "" });


                var p = doc.InsertParagraph();

                p.InsertTableBeforeSelf(AddTable(datas));
                var h1 = doc.InsertParagraph("測試加入文字");
                h1.StyleId = "Heading1";
                h1.InsertTableAfterSelf(AddTable(datas));

                doc.Save();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (WordprocessingDocument package = WordprocessingDocument.Open("D:\\Test.docx", true))
            {
                WaterMarkService.InsertCustomWatermark(package, "D:\\下載.jpg");
                package.SaveAs("D:\\Test1.docx").Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Word.Document wordDocument;
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            wordDocument = word.Documents.Open("D:\\Test.docx");
            wordDocument.TablesOfContents[1].Update();
            wordDocument.ExportAsFixedFormat("D:\\Test.pdf", WdExportFormat.wdExportFormatPDF);
            wordDocument.Close();
            word.Quit();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AddWatermark("D:\\Test.pdf", "D:\\Test_Watermark.pdf", "D:\\TKU-logo-12cm.jpg");
        }

        public void Watermark(string inputPath, string outputPath, string watermarkPath)
        {
            try
            {
                PdfReader pdfReader = new PdfReader(inputPath);
                int numberOfPages = pdfReader.NumberOfPages;
                FileStream outputStream = new FileStream(outputPath, FileMode.Create);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, outputStream);
                PdfContentByte waterMarkContent;

                string watermarkimagepath = watermarkPath;
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(watermarkimagepath);

                image.SetAbsolutePosition(255, 785);
                for (int i = 1; i <= numberOfPages; i++)
                {
                    waterMarkContent = pdfStamper.GetUnderContent(i);
                    waterMarkContent.AddImage(image);
                }
                pdfStamper.Close();
                pdfReader.Close();
            }
            catch (Exception ex)
            {                
                throw ex;
            }
        }
        public bool AddWatermark(string inputPath, string outputPath, string watermarkPath)
        {
            try
            {
                PdfReader pdfReader = new PdfReader(inputPath);
                int numberOfPages = pdfReader.NumberOfPages;
                FileStream outputStream = new FileStream(outputPath, FileMode.Create);
                PdfStamper pdfStamper = new PdfStamper(pdfReader, outputStream);
                PdfContentByte waterMarkContent;

                iTextSharp.text.Image image = null;
                if (string.IsNullOrEmpty(watermarkPath))
                {
                    Stream s = GetType().Assembly.GetManifestResourceStream("WatermarkTool.wm.png");
                    image = iTextSharp.text.Image.GetInstance(s);
                }
                else
                {
                    image = iTextSharp.text.Image.GetInstance(watermarkPath);
                }

                // A4尺寸 595x842    圖片最左下角為(0,0)點
                float imagesize_X = 159;
                float imagesize_Y = 159;
                float page_Size_X = 595;
                float page_Size_Y = 842;

                PdfGState pdfgstate = new PdfGState()
                {
                    FillOpacity = 0.4f,
                    StrokeOpacity = 0.4f
                };

                image.ScaleAbsolute(imagesize_X, imagesize_X);
                image.SetAbsolutePosition(((page_Size_X- imagesize_X)/2), ((page_Size_Y - imagesize_Y) / 2));    /*計算中心點*/
                
                for (int i = 1; i <= numberOfPages; i++)
                {
                    waterMarkContent = pdfStamper.GetUnderContent(i);
                    waterMarkContent.SetGState(pdfgstate); //寫入入設定的透明度
                    waterMarkContent.AddImage(image);
                }
                pdfStamper.Close();
                pdfReader.Close();
                outputStream.Close();
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

    }
    public class testdata
    {
        public string Step { get; set; }
        public string StepDesc { get; set; }
        public string InvolvedPost { get; set; }
        public string Remark { get; set; }
    }
}
