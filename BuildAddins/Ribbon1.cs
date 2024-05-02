using Microsoft.Office.Core;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BuildAddins
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        //áp dụng kiểu
        private bool ApplyCharStyle(Style objStyle)
        {
            // Applies a character style to the current selection
            // If there is no highlighted selection, expand it until the next space or paragraph is found

            try
            {
                if (objStyle == null || objStyle.Type != WdStyleType.wdStyleTypeCharacter)
                    return false;

                Selection selection = Globals.ThisAddIn.Application.Selection;

                // if no text is highlighted, expand the selection up to the next space or paragraph
                if (selection.Start == selection.End)
                {
                    selection.MoveStartUntil(" " + "\r", WdConstants.wdBackward);
                    selection.MoveEndUntil(" " + "\r", WdConstants.wdForward);
                }

                selection.set_Style(objStyle);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("An error occurred in the macro code (ApplyCharStyle): " + ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        private bool ApplyParaStyle(Style objStyle, bool booMultiPara)
        {
            // Applies a paragraph style to the current selection
            // If the booMultiPara flag is not active, the function is cancelled for multi-paragraph selections
            // Set the cursor to the beginning of the current paragraph

            try
            {
                if (objStyle == null || objStyle.Type != WdStyleType.wdStyleTypeParagraph)
                    return false;

                Selection selection = Globals.ThisAddIn.Application.Selection;

                // check whether text is highlighted
                if (selection.Start != selection.End)
                {
                    // some text is selected
                    if (selection.End > selection.Paragraphs[1].Range.End)
                    {
                        // multiple paragraphs are selected
                        if (!booMultiPara)
                        {
                            // if not supported, cancel
                            System.Windows.Forms.MessageBox.Show("This function is not available if more than one paragraph is selected!", "Warning", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Exclamation);
                            return false;
                        }
                    }
                }

                selection.ParagraphFormat.set_Style(objStyle);
                // collapse the selection
                selection.Collapse(WdCollapseDirection.wdCollapseStart);
                // go up, if the cursor is not at the beginning of the paragraph
                if (selection.Start > selection.Paragraphs[1].Range.Start)
                {
                    selection.MoveUp(WdUnits.wdParagraph, 1);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Lỗi xuất hiện tại đoạn mã (ApplyParaStyle): " + ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
        private Style GetOrCreateStyle(string styleName, string fontName, int fontSize, bool isBold, bool isItalic, WdParagraphAlignment alignment, WdLineSpacing lineSpacingRule, float leftIndent)
        {
            try
            {
                Style style = Globals.ThisAddIn.Application.ActiveDocument.Styles[styleName];

                if (style == null)
                {
                    style = Globals.ThisAddIn.Application.ActiveDocument.Styles.Add(styleName);
                    style.Font.Name = fontName;
                    style.Font.Size = fontSize;
                    style.Font.Bold = isBold ? 1 : 0;
                    style.Font.Italic = isItalic ? 1 : 0;
                    style.ParagraphFormat.Alignment = alignment;
                    style.ParagraphFormat.LineSpacingRule = lineSpacingRule;
                    style.ParagraphFormat.FirstLineIndent = leftIndent;
                }

                return style;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Lỗi tại đoạn mã  (GetOrCreateStyle): " + ex.Message, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return null;
            }
        }

        private void btnInsertImage_Click(object sender, RibbonControlEventArgs e)
        {

            try
            {
                // Open file dialog to select image
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Image Files (*.bmp;*.jpg;*.jpeg;*.gif;*.png)|*.BMP;*.JPG;*.JPEG;*.GIF;*.PNG|All files (*.*)|*.*";
                openFileDialog.Title = "Select an image file";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Get selected image path
                    string imagePath = openFileDialog.FileName;

                    // Insert image into Word document
                    Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
                    wordApp.ActiveDocument.InlineShapes.AddPicture(imagePath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗ Tại: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

        private void btnTenBang_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
                Microsoft.Office.Interop.Word.Document doc = wordApp.ActiveDocument;

                if (doc.Tables.Count == 0)
                {
                    MessageBox.Show("Không có bảng nào trong tài liệu.", "Không có bảng", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Microsoft.Office.Interop.Word.Table selectedTable = null;
                foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
                {
                    if (wordApp.Selection.Range.InRange(table.Range))
                    {
                        selectedTable = table;
                        break;
                    }
                }

                if (selectedTable == null)
                {
                    MessageBox.Show("Vui lòng chọn hoặc nhấn vào 1 bảng.", "Chưa có bảng được chọn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int tableIndex = 1;
                foreach (Microsoft.Office.Interop.Word.Table table in doc.Tables)
                {
                    if (table == selectedTable)
                        break;
                    tableIndex++;
                }

                string tableName = "Bảng " + tableIndex.ToString();

                // Lấy dòng văn bản trước bảng đã chọn
                Range previousParagraphRange = selectedTable.Range.Previous(Microsoft.Office.Interop.Word.WdUnits.wdParagraph, 1);

                // Nếu không có dòng văn bản trước, không thể chèn ngoài bảng, bạn có thể xử lý tùy theo yêu cầu của bạn
                if (previousParagraphRange != null)
                {
                    // Chèn tên bảng vào trước dòng văn bản trước bảng đã chọn
                    previousParagraphRange.InsertBefore(tableName+". Bạn nhập tên bảng tại đây!");
                }
                else 
                {
                    selectedTable.Range.Text = tableName;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("An error occurred: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnTenAnh_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
                Document doc = wordApp.ActiveDocument;

                // Kiểm tra xem người dùng đã chọn một hình ảnh hay không
                if (wordApp.Selection.Type == WdSelectionType.wdSelectionInlineShape)
                {
                    InlineShape selectedInlineShape = wordApp.Selection.InlineShapes[1]; // Lấy hình ảnh được chọn

                    // Kiểm tra nếu hình ảnh là loại Picture
                    if (selectedInlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        int imageIndex = 0;
                        foreach (InlineShape inlineShape in doc.InlineShapes)
                        {
                            imageIndex++;
                            if (inlineShape == selectedInlineShape)
                            {
                                // Tạo tên cho hình ảnh
                                string imageName = "Hình Ảnh " + imageIndex;

                                // Tạo một dòng mới trước hình ảnh
                                InlineShape newParagraph = selectedInlineShape.Range.Paragraphs.Add().Range.InlineShapes.AddOLEObject();
                                newParagraph.Range.Text = imageName;
                                newParagraph.Title = imageName;                           
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Hãy chọn một hình ảnh.", "Không có hình ảnh được chọn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Hãy chọn một hình ảnh.", "Không có hình ảnh được chọn", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button23_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application wordApp = Globals.ThisAddIn.Application;
                Document doc = wordApp.ActiveDocument;

                int imageCount = 0;

                foreach (InlineShape inlineShape in doc.InlineShapes)
                {
                    if (inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        imageCount++;
                    }
                }

                MessageBox.Show("Số lượng hình ảnh trong tài liệu: " + imageCount.ToString(), "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
