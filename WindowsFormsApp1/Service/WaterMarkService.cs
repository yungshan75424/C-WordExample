using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using V = DocumentFormat.OpenXml.Vml;
using Ovml = DocumentFormat.OpenXml.Vml.Office;
using Wvml = DocumentFormat.OpenXml.Vml.Wordprocessing;
using DocumentFormat.OpenXml;

namespace WindowsFormsApp1.Service
{
    class WaterMarkService
    {
        private static string imagePart1Data = "";

        public static void InsertCustomWatermark(WordprocessingDocument package, string p)
        {
            SetWaterMarkPicture(p);
            MainDocumentPart mainDocumentPart1 = package.MainDocumentPart;
            if (mainDocumentPart1 != null)
            {
                mainDocumentPart1.DeleteParts(mainDocumentPart1.HeaderParts);
                HeaderPart headPart1 = mainDocumentPart1.AddNewPart<HeaderPart>();
                GenerateHeaderPart1Content(headPart1);
                string rId = mainDocumentPart1.GetIdOfPart(headPart1);
                ImagePart image = headPart1.AddNewPart<ImagePart>("image/jpeg", "rId1");
                GenerateImagePart1Content(image);
                IEnumerable<SectionProperties> sectPrs = mainDocumentPart1.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    sectPr.RemoveAllChildren<HeaderReference>();
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
                }
            }
            else
            {
                MessageBox.Show("alert");
            }
        }
        private static void GenerateHeaderPart1Content(HeaderPart headerPart1)
        {
            Header header1 = new Header();
            Paragraph paragraph2 = new Paragraph();
            Run run1 = new Run();
            Picture picture1 = new Picture();
            V.Shape shape1 = new V.Shape() { Id = "WordPictureWatermark75517470", Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:415.2pt;height:456.15pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", OptionalString = "_x0000_s2051", AllowInCell = false, Type = "#_x0000_t75" };
            V.ImageData imageData1 = new V.ImageData() { Gain = "19661f", BlackLevel = "22938f", Title = "水印", RelationshipId = "rId1" };
            shape1.Append(imageData1);
            picture1.Append(shape1);
            run1.Append(picture1);
            paragraph2.Append(run1);
            header1.Append(paragraph2);
            headerPart1.Header = header1;
        }

        private static void GenerateHeaderPart1Content1(HeaderPart headerPart1)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 w15 w16se wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("cx", "http://schemas.microsoft.com/office/drawing/2014/chartex");
            header1.AddNamespaceDeclaration("cx1", "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex");
            header1.AddNamespaceDeclaration("cx2", "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("w15", "http://schemas.microsoft.com/office/word/2012/wordml");
            header1.AddNamespaceDeclaration("w16se", "http://schemas.microsoft.com/office/word/2015/wordml/symex");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph2 = new Paragraph() { RsidParagraphAddition = "00936B9A", RsidRunAdditionDefault = "00936B9A" };

            ParagraphProperties paragraphProperties2 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId2 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties2.Append(paragraphStyleId2);
            BookmarkStart bookmarkStart1 = new BookmarkStart() { Name = "_GoBack", Id = "0" };

            Run run1 = new Run();

            RunProperties runProperties1 = new RunProperties();
            NoProof noProof1 = new NoProof();

            runProperties1.Append(noProof1);

            Picture picture1 = new Picture();

            V.Shapetype shapetype1 = new V.Shapetype() { Id = "_x0000_t136", CoordinateSize = "21600,21600", OptionalNumber = 136, Adjustment = "10800", EdgePath = "m@7,l@8,m@5,21600l@6,21600e" };

            V.Formulas formulas1 = new V.Formulas();
            V.Formula formula1 = new V.Formula() { Equation = "sum #0 0 10800" };
            V.Formula formula2 = new V.Formula() { Equation = "prod #0 2 1" };
            V.Formula formula3 = new V.Formula() { Equation = "sum 21600 0 @1" };
            V.Formula formula4 = new V.Formula() { Equation = "sum 0 0 @2" };
            V.Formula formula5 = new V.Formula() { Equation = "sum 21600 0 @3" };
            V.Formula formula6 = new V.Formula() { Equation = "if @0 @3 0" };
            V.Formula formula7 = new V.Formula() { Equation = "if @0 21600 @1" };
            V.Formula formula8 = new V.Formula() { Equation = "if @0 0 @2" };
            V.Formula formula9 = new V.Formula() { Equation = "if @0 @4 21600" };
            V.Formula formula10 = new V.Formula() { Equation = "mid @5 @6" };
            V.Formula formula11 = new V.Formula() { Equation = "mid @8 @5" };
            V.Formula formula12 = new V.Formula() { Equation = "mid @7 @8" };
            V.Formula formula13 = new V.Formula() { Equation = "mid @6 @7" };
            V.Formula formula14 = new V.Formula() { Equation = "sum @6 0 @5" };

            formulas1.Append(formula1);
            formulas1.Append(formula2);
            formulas1.Append(formula3);
            formulas1.Append(formula4);
            formulas1.Append(formula5);
            formulas1.Append(formula6);
            formulas1.Append(formula7);
            formulas1.Append(formula8);
            formulas1.Append(formula9);
            formulas1.Append(formula10);
            formulas1.Append(formula11);
            formulas1.Append(formula12);
            formulas1.Append(formula13);
            formulas1.Append(formula14);
            V.Path path1 = new V.Path() { AllowTextPath = true, ConnectionPointType = Ovml.ConnectValues.Custom, ConnectionPoints = "@9,0;@10,10800;@11,21600;@12,10800", ConnectAngles = "270,180,90,0" };
            V.TextPath textPath1 = new V.TextPath() { On = true, FitShape = true };

            V.ShapeHandles shapeHandles1 = new V.ShapeHandles();
            V.ShapeHandle shapeHandle1 = new V.ShapeHandle() { Position = "#0,bottomRight", XRange = "6629,14971" };

            shapeHandles1.Append(shapeHandle1);
            Ovml.Lock lock1 = new Ovml.Lock() { Extension = V.ExtensionHandlingBehaviorValues.Edit, TextLock = true, ShapeType = true };

            shapetype1.Append(formulas1);
            shapetype1.Append(path1);
            shapetype1.Append(textPath1);
            shapetype1.Append(shapeHandles1);
            shapetype1.Append(lock1);

            V.Shape shape1 = new V.Shape() { Id = "_x0000_s2049", Style = "margin-margin-width:87.75pt;height:51pt;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin", Type = "#_x0000_t136" };
            V.Fill fill1 = new V.Fill() { Title = "", RelationshipId = "rId1" };
            V.Stroke stroke1 = new V.Stroke() { Title = "", RelationshipId = "rId1" };
            V.Shadow shadow1 = new V.Shadow() { Color = "#868686" };
            V.TextPath textPath2 = new V.TextPath() { Style = "font-family:\"Arial Black\";v-text-kern:t", FitPath = true, Trim = true, String = "Test" };
            Wvml.TextWrap textWrap1 = new Wvml.TextWrap() { Type = Wvml.WrapValues.Square, AnchorX = Wvml.HorizontalAnchorValues.Margin, AnchorY = Wvml.VerticalAnchorValues.Margin };

            shape1.Append(fill1);
            shape1.Append(stroke1);
            shape1.Append(shadow1);
            shape1.Append(textPath2);
            shape1.Append(textWrap1);

            picture1.Append(shapetype1);
            picture1.Append(shape1);

            run1.Append(runProperties1);
            run1.Append(picture1);
            BookmarkEnd bookmarkEnd1 = new BookmarkEnd() { Id = "0" };

            paragraph2.Append(paragraphProperties2);
            paragraph2.Append(bookmarkStart1);
            paragraph2.Append(run1);
            paragraph2.Append(bookmarkEnd1);

            header1.Append(paragraph2);

            headerPart1.Header = header1;
        }


        private static void GenerateImagePart1Content(ImagePart imagePart1)
        {
            System.IO.Stream data = GetBinaryDataStream(imagePart1Data);
            imagePart1.FeedData(data);
            data.Close();
        }

        

        private static System.IO.Stream GetBinaryDataStream(string base64String)
        {
            return new System.IO.MemoryStream(System.Convert.FromBase64String(base64String));
        }
        public static void SetWaterMarkPicture(string file)
        {
            FileStream inFile;
            byte[] byteArray;
            try
            {
                inFile = new FileStream(file, FileMode.Open, FileAccess.Read);
                byteArray = new byte[inFile.Length];
                long byteRead = inFile.Read(byteArray, 0, (int)inFile.Length);
                inFile.Close();
                imagePart1Data = Convert.ToBase64String(byteArray, 0, byteArray.Length);
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message);
            }
        }
    }
}

