using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using Microsoft.Office.Interop.Word;
using System.Diagnostics.Eventing.Reader;
using System.Windows.Forms;

namespace BuildAddins
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            PageSetUp();

        }
        public void PageSetUp()
        {
           

                Word.Document doc = null;

                if (Application.Documents.Count > 0)
                {
                    doc = Application.ActiveDocument;
                }
                else
                {
                // doc = Application.Documents.Add();
                }

                if (doc != null)
                {
                    // Thiết lập thông số PageSetup
                    doc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                    doc.PageSetup.PaperSize = WdPaperSize.wdPaperA4;
                    doc.PageSetup.TopMargin = ConvertCentimetersToPoints(2);
                    doc.PageSetup.BottomMargin = ConvertCentimetersToPoints(2);
                    doc.PageSetup.LeftMargin = ConvertCentimetersToPoints(3);
                    doc.PageSetup.RightMargin = ConvertCentimetersToPoints(2);
                    doc.PageSetup.HeaderDistance = ConvertCentimetersToPoints(2.85f);
                    doc.PageSetup.FooterDistance = ConvertCentimetersToPoints(2.85f);

                    // Thiết lập định dạng văn bản cho nội dung bài báo
                    doc.Content.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    doc.Content.ParagraphFormat.SpaceAfter = 0;
                    doc.Content.ParagraphFormat.SpaceBefore = 0;
                    doc.Content.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle;
                    doc.Content.ParagraphFormat.LeftIndent = ConvertCentimetersToPoints(0.5f);

                    // Thiết lập định dạng cho tiêu đề các phần
                    foreach (Section section in doc.Sections)
                    {
                        section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                        section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }

                    // Thiết lập khoảng trắng dòng cho tiêu đề
                    foreach (Paragraph titleParagraph in doc.Paragraphs)
                    {
                        if (
                            titleParagraph.get_Style() == "Heading 1" || titleParagraph.get_Style() == "Heading 2")
                        {
                            titleParagraph.Format.SpaceBefore = 6;
                            titleParagraph.Format.SpaceAfter = 6;
                            titleParagraph.Format.LeftIndent = 0;
                        }
                    }

                    // Thiết lập font chữ và cỡ chữ cho toàn bộ nội dung
                    Font contentfont = doc.Content.Font;
                    doc.Content.Font.Name = "Times New Roman";
                    doc.Content.Font.Size = 12;
                }
                else
                {
                    Console.WriteLine("Không thể truy cập tài liệu!");
                }
           
        }

    
        private float ConvertCentimetersToPoints(float centimeters)
        {
            // 1 centimeter = 28.35 points
            return centimeters * 28.35f;
        }
    

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
