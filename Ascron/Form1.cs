using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace EmailToPdfConverter
{
    public partial class Form1 : Form
    {
        private readonly Outlook.Application outlookApp;
        private readonly Outlook.NameSpace outlookNamespace;

        // Preset paths for the queues
        private readonly string queuePath1 = @"\\csputascron01\DE_Sources\Queue1";
        private readonly string queuePath2 = @"\\csputascron01\DE_Sources\Queue2";
        private readonly string queuePath3 = @"\\csputascron01\DE_Sources\Queue3";
        private readonly string queuePath4 = @"\\csputascron01\DE_Sources\Queue4";
        private readonly string queuePath5 = @"\\csputascron01\DE_Sources\Queue5";
        private readonly string queuePath6 = @"C:\Users\de-Angelo\OneDrive - PRA Group Europe\Dokumente\New folder (10)";

        private string previewFilePath = string.Empty;
        private WebBrowser pdfPreviewBrowser;

        public Form1()
        {
            InitializeComponent();

            // Set form style
            this.BackColor = System.Drawing.Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Initialize Outlook
            try
            {
                outlookApp = new Outlook.Application();
                outlookNamespace = outlookApp.GetNamespace("MAPI");
                outlookNamespace.Logon("", "", Missing.Value, Missing.Value);
            }
            catch (COMException ex) when (ex.ErrorCode == -2147023174) // 0x800706ba
            {
                MessageBox.Show($"Error initializing Outlook: RPC server is unavailable. Make sure Outlook is running and connected. {ex.Message}");
                throw;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing Outlook: {ex.Message}");
                throw;
            }

            // Initialize PDF Preview Browser
            pdfPreviewBrowser = new WebBrowser
            {
                Location = new System.Drawing.Point(12, 70),
                MinimumSize = new System.Drawing.Size(20, 20),
                Size = new System.Drawing.Size(600, 400)
            };
            this.Controls.Add(pdfPreviewBrowser);
        }

        private void btnMakePreview_Click(object sender, EventArgs e)
        {
            var openMail = GetCurrentOpenEmail();
            if (openMail != null)
            {
                previewFilePath = ConvertEmailToPdf(openMail);
                if (!string.IsNullOrEmpty(previewFilePath))
                {
                    ShowPdfPreview(previewFilePath);
                }
            }
        }

        private void btnQueue1_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath1);
        private void btnQueue2_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath2);
        private void btnQueue3_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath3);
        private void btnQueue4_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath4);
        private void btnQueue5_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath5);
        private void btnQueue6_Click(object sender, EventArgs e) => SendPreviewToQueue(queuePath6);

        private void chkTopMost_CheckedChanged(object sender, EventArgs e)
        {
            this.TopMost = chkTopMost.Checked;
        }

        private Outlook.MailItem GetCurrentOpenEmail()
        {
            try
            {
                var inspector = outlookApp.ActiveInspector();
                if (inspector != null && inspector.CurrentItem is Outlook.MailItem mailItem)
                {
                    return mailItem;
                }
                else
                {
                    MessageBox.Show("No email is currently open in Outlook.");
                }
            }
            catch (COMException ex) when (ex.ErrorCode == -2147023174) // 0x800706ba
            {
                MessageBox.Show($"Error retrieving the current open email: RPC server is unavailable. Make sure Outlook is running and connected. {ex.Message}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error retrieving the current open email: {ex.Message}");
            }
            return null;
        }

        private string ConvertEmailToPdf(Outlook.MailItem mail)
        {
            string subject = mail.Subject;
            if (string.IsNullOrEmpty(subject)) return string.Empty;

            Log("ConvertEmailToPdf", subject, "Started");

            subject = new string(subject.Where(c => !Path.GetInvalidFileNameChars().Contains(c)).ToArray());
            string tempFolderPath = Path.Combine(Path.GetTempPath(), "EmailToPdfConverter");

            if (!Directory.Exists(tempFolderPath))
            {
                Directory.CreateDirectory(tempFolderPath);
                Log("DirectoryCreation", tempFolderPath, "Success");
            }

            string tempMailFilePath = Path.Combine(tempFolderPath, subject + ".pdf");

            // Save email as .mht and convert to PDF
            string mhtFilePath = Path.Combine(tempFolderPath, subject + ".mht");
            if (!SaveMailAsMht(mail, mhtFilePath))
            {
                Log("SaveMailAsMht", mhtFilePath, "Failure", "Failed to save email as MHT.");
                return string.Empty;
            }

            try
            {
                string emailPdfPath = Path.Combine(tempFolderPath, subject + "_email.pdf");
                ConvertMhtToPdf(mhtFilePath, emailPdfPath);
                File.Delete(mhtFilePath);
                Log("ConvertMhtToPdf", emailPdfPath, "Success");

                // Convert attachments
                var attachmentPdfPaths = ConvertAttachmentsToPdf(mail, tempFolderPath);
                Log("ConvertAttachmentsToPdf", string.Join(", ", attachmentPdfPaths), "Success");

                // Merge all PDFs
                MergePdfs(emailPdfPath, attachmentPdfPaths, tempMailFilePath);
                Log("MergePdfs", tempMailFilePath, "Success");

                // Cleanup
                File.Delete(emailPdfPath);
                foreach (var path in attachmentPdfPaths)
                {
                    File.Delete(path);
                    Log("Cleanup", path, "Deleted");
                }

                MessageBox.Show("Email and attachments converted to PDF successfully.");
                Log("ConvertEmailToPdf", tempMailFilePath, "Completed");
                return tempMailFilePath;
            }
            catch (Exception ex)
            {
                Log("ConvertEmailToPdf", tempMailFilePath, "Failure", ex.Message);
                return string.Empty;
            }
        }
        private void Log(string action, string fileName, string status, string message = "")
        {
            string appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "EmailToPdfConverter");
            if (!Directory.Exists(appDataPath))
            {
                Directory.CreateDirectory(appDataPath);
            }
            string logFilePath = Path.Combine(appDataPath, "file_operations.log");
            string logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - ACTION: {action}, FILE: {fileName}, STATUS: {status}, MESSAGE: {message}";

            try
            {
                File.AppendAllText(logFilePath, logMessage + Environment.NewLine);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to write to log file: {ex.Message}");
            }
        }

        private void SendPreviewToQueue(string queuePath)
        {
            if (string.IsNullOrEmpty(previewFilePath))
            {
                MessageBox.Show("Please generate a preview first.");
                Log("SendPreviewToQueue", "NoFile", "Failure", "No preview file generated.");
                return;
            }

            string finalMailFilePath = Path.Combine(queuePath, Path.GetFileName(previewFilePath));
            try
            {
                // Dispose and reinitialize WebBrowser to release the file
                pdfPreviewBrowser.Dispose();
                pdfPreviewBrowser = new WebBrowser
                {
                    Location = new System.Drawing.Point(12, 70),
                    MinimumSize = new System.Drawing.Size(20, 20),
                    Size = new System.Drawing.Size(600, 400)
                };
                this.Controls.Add(pdfPreviewBrowser);

                // Delay to ensure the file is not being used
                Thread.Sleep(1000);

                File.Move(previewFilePath, finalMailFilePath);
                Log("SendPreviewToQueue", finalMailFilePath, "Success");
                MessageBox.Show("PDF successfully moved to queue.");
                previewFilePath = string.Empty; // Reset the preview file path
            }
            catch (IOException ex)
            {
                if (ex.HResult == 0x800700B7) // ERROR_ALREADY_EXISTS
                {
                    Log("SendPreviewToQueue", finalMailFilePath, "Failure", "File already exists.");
                    MessageBox.Show($"The file '{finalMailFilePath}' already exists.");
                }
                else
                {
                    Log("SendPreviewToQueue", finalMailFilePath, "Failure", ex.Message);
                    MessageBox.Show($"An error occurred while moving the file: {ex.Message}");
                }
            }
        }

        private static bool SaveMailAsMht(Outlook.MailItem mail, string mhtFilePath)
        {
            int maxRetries = 3;
            int delay = 2000; // milliseconds
            for (int attempt = 0; attempt < maxRetries; attempt++)
            {
                try
                {
                    mail.SaveAs(mhtFilePath, Outlook.OlSaveAsType.olMHTML);
                    return true;
                }
                catch (COMException ex) when (ex.ErrorCode == -2147467259) // 0x80004005
                {
                    if (attempt == maxRetries - 1)
                    {
                        MessageBox.Show($"Failed to save email as .mht after {maxRetries} attempts: {ex.Message}");
                        return false;
                    }
                    Thread.Sleep(delay); // Wait before retrying
                }
            }
            return false;
        }

        private void ConvertMhtToPdf(string mhtFilePath, string pdfFilePath)
        {
            var wordApp = new Word.Application();
            var documents = wordApp.Documents;
            var doc = documents.Open(mhtFilePath);

            // Loop through all inline shapes (images, etc.) in the document and adjust their size
            foreach (Word.InlineShape inlineShape in doc.InlineShapes)
            {
                if (inlineShape.Type == Word.WdInlineShapeType.wdInlineShapePicture ||
                    inlineShape.Type == Word.WdInlineShapeType.wdInlineShapeLinkedPicture)
                {
                    // Lock aspect ratio
                    inlineShape.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    float originalWidth = inlineShape.Width;
                    float originalHeight = inlineShape.Height;

                    // Adjust size based on page width and height, keeping the aspect ratio
                    if (originalWidth > doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin)
                    {
                        float scaleFactor = (doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin) / originalWidth;
                        inlineShape.Width = originalWidth * scaleFactor;
                        inlineShape.Height = originalHeight * scaleFactor;
                    }

                    // Ensure small images don't get stretched
                    if (inlineShape.Width < 100)
                    {
                        inlineShape.Width = originalWidth; // Revert any scaling for small images
                        inlineShape.Height = originalHeight;
                    }
                }
            }

            doc.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatPDF);
            doc.Close(false);
            wordApp.Quit();
            Marshal.ReleaseComObject(doc);
            Marshal.ReleaseComObject(documents);
            Marshal.ReleaseComObject(wordApp);
        }

        private static string[] ConvertAttachmentsToPdf(Outlook.MailItem mail, string folderPath)
        {
            var attachmentPdfPaths = new List<string>();

            foreach (Outlook.Attachment attachment in mail.Attachments)
            {
                string attachmentFileName = attachment.FileName;
                if (string.IsNullOrEmpty(attachmentFileName)) continue;

                string attachmentPath = Path.Combine(folderPath, attachmentFileName);
                attachment.SaveAsFile(attachmentPath);

                if (attachmentPath.EndsWith(".doc") || attachmentPath.EndsWith(".docx"))
                {
                    string pdfPath = ConvertWordToPdf(attachmentPath);
                    attachmentPdfPaths.Add(pdfPath);
                }
                else if (attachmentPath.EndsWith(".xls") || attachmentPath.EndsWith(".xlsx"))
                {
                    string pdfPath = ConvertExcelToPdf(attachmentPath);
                    attachmentPdfPaths.Add(pdfPath);
                }
                else if (attachmentPath.EndsWith(".pdf"))
                {
                    attachmentPdfPaths.Add(attachmentPath); // Already a PDF
                }
                else if (attachmentPath.EndsWith(".jpg") || attachmentPath.EndsWith(".jpeg") || attachmentPath.EndsWith(".png"))
                {
                    string pdfPath = ConvertImageToPdf(attachmentPath);
                    attachmentPdfPaths.Add(pdfPath);
                }
                else
                {
                    // Unsupported file type, delete it
                    File.Delete(attachmentPath);
                }
            }

            return attachmentPdfPaths.ToArray();
        }

        private static string ConvertWordToPdf(string wordFilePath)
        {
            var wordApp = new Word.Application();
            var documents = wordApp.Documents;
            var doc = documents.Open(wordFilePath);

            string pdfFilePath = Path.ChangeExtension(wordFilePath, ".pdf");
            doc.SaveAs2(pdfFilePath, Word.WdSaveFormat.wdFormatPDF);
            doc.Close();
            wordApp.Quit();
            File.Delete(wordFilePath);
            Marshal.ReleaseComObject(doc);
            Marshal.ReleaseComObject(documents);
            Marshal.ReleaseComObject(wordApp);
            return pdfFilePath;
        }

        private static string ConvertExcelToPdf(string excelFilePath)
        {
            var excelApp = new Excel.Application();
            var workbooks = excelApp.Workbooks;
            var wb = workbooks.Open(excelFilePath);

            string pdfFilePath = Path.ChangeExtension(excelFilePath, ".pdf");
            wb.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
            wb.Close();
            excelApp.Quit();
            File.Delete(excelFilePath);
            Marshal.ReleaseComObject(wb);
            Marshal.ReleaseComObject(workbooks);
            Marshal.ReleaseComObject(excelApp);
            return pdfFilePath;
        }

        private static string ConvertImageToPdf(string imagePath)
        {
            string pdfPath = Path.ChangeExtension(imagePath, ".pdf");
            using (var document = new Document())
            {
                using (var stream = new FileStream(pdfPath, FileMode.Create))
                {
                    PdfWriter.GetInstance(document, stream);
                    document.Open();
                    using (var imageStream = new FileStream(imagePath, FileMode.Open))
                    {
                        var image = iTextSharp.text.Image.GetInstance(imageStream);
                        float pageWidth = document.PageSize.Width - document.LeftMargin - document.RightMargin;
                        float pageHeight = document.PageSize.Height - document.TopMargin - document.BottomMargin;

                        if (image.Width > pageWidth || image.Height > pageHeight)
                        {
                            image.ScaleToFit(pageWidth, pageHeight);
                        }
                        else
                        {
                            image.ScaleToFit(image.Width, image.Height);
                        }
                        image.Alignment = Element.ALIGN_CENTER;
                        document.Add(image);
                    }
                    document.Close();
                }
            }
            File.Delete(imagePath);
            return pdfPath;
        }

        private static void MergePdfs(string mainPdf, string[] attachmentPdfs, string outputPdf)
        {
            using (var stream = new FileStream(outputPdf, FileMode.Create))
            {
                var document = new Document();
                var pdfCopy = new PdfCopy(document, stream);
                document.Open();

                AddPdfToCopy(pdfCopy, mainPdf);

                foreach (var attachmentPdf in attachmentPdfs)
                {
                    AddPdfToCopy(pdfCopy, attachmentPdf);
                }

                document.Close();
            }
        }

        private static void AddPdfToCopy(PdfCopy pdfCopy, string pdfPath)
        {
            using (var reader = new PdfReader(pdfPath))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    pdfCopy.AddPage(pdfCopy.GetImportedPage(reader, i));
                }
            }
        }

        private void ShowPdfPreview(string pdfPath)
        {
            pdfPreviewBrowser.Navigate(pdfPath);
        }
    }
}
