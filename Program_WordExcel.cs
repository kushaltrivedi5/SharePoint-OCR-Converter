using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using Microsoft.SharePoint.Client;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf;
using iText.Pdfocr;
using iText.Pdfocr.Tesseract4;
using iText.License;

namespace ReadPdfFromEventReciever
{
    class Program
    {
        static string webSPOUrl = "https://startvindia.sharepoint.com/sites/uathr";
        private static readonly Tesseract4OcrEngineProperties tesseract4OcrEngineProperties = new Tesseract4OcrEngineProperties();
       
        static void Main(string[] args)
        {
            try
            {
                OfficeDevPnP.Core.AuthenticationManager authManager = new OfficeDevPnP.Core.AuthenticationManager();

                var cxt = authManager.GetWebLoginClientContext(webSPOUrl);
                {
                    Web web = cxt.Web;
                    User user = web.CurrentUser;
                    cxt.Load(web);
                    cxt.Load(user);
                    cxt.ExecuteQuery();
                    cxt.Load(web.Lists,
                        lists => lists.Include(list => list.Title,
                            list => list.Id));
                    cxt.ExecuteQuery();
                    Console.ForegroundColor = ConsoleColor.White;



                    //var codefound = ReadPdfFile(cxt);
                    //var codefound = ReadWordFile(cxt);
                    var codefound = convertToOCR();
                    if (codefound == true)
                    {
                        //PDF APPROVED CODE GOES HERE
                    }
                    Console.WriteLine(codefound);
                    Console.Read();
                }

            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(string.Format("Exception {0}.", ex.ToString()));
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("Press enter to continue.");
                Console.Read();
            }
        }
        public static Boolean ReadPdfFile(ClientContext cxt)
        {
            List mdl = cxt.Web.Lists.GetByTitle("Documents");
            CamlQuery q = new CamlQuery();
            string docname = "testocr.pdf";
            q.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>" + docname + "</Value></Eq></Where></Query></View>";

            ListItemCollection allfile = mdl.GetItems(q);
            cxt.Load(mdl);
            cxt.Load(allfile);
            cxt.ExecuteQuery();

            Boolean codeExist = false;

            foreach (Microsoft.SharePoint.Client.ListItem pdf in allfile)
            {
                string fileName = webSPOUrl + "/Shared%20Documents/" + Convert.ToString(pdf.FieldValues["FileLeafRef"]);
                var file = mdl.RootFolder.Files.GetByUrl(fileName);
                cxt.Load(file);
                cxt.ExecuteQuery();

                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                cxt.Load(file);
                cxt.ExecuteQuery();

                string textPDF = string.Empty;

                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        /*data.Value.CopyTo(mStream);
                        PdfReader reader = new PdfReader(mStream);
                        var pdfFile = new PdfDocument(reader);*/

                        IList<FileInfo> LIST_IMAGES_OCR = new List<FileInfo>();
                            
                        LIST_IMAGES_OCR.Add(new FileInfo(@"../1.png"));
                        LIST_IMAGES_OCR.Add(new FileInfo(@"../2.png"));
                        LIST_IMAGES_OCR.Add(new FileInfo(@"../3.png"));

                        var tesseractReader = new Tesseract4LibOcrEngine(tesseract4OcrEngineProperties);
                        tesseract4OcrEngineProperties.SetPathToTessData(new FileInfo(@"C:\Users\kushal.trivedi.vc\source\repos\PDFtoOCR\bin\TESS_DATA"));

                        var properties = new OcrPdfCreatorProperties();
                        properties.SetPdfLang("en");
                        var ocrPdfCreator = new OcrPdfCreator(tesseractReader, properties);
                        var writer = new PdfWriter(new FileInfo(@"C:\Users\kushal.trivedi.vc\source\repos\PDFtoOCR\bin\PDF_Output\NewPdf.pdf"));
                        
                        ocrPdfCreator.CreatePdf(LIST_IMAGES_OCR, writer).Close();
                        

                    }
                }

            }
            return codeExist;
        }

       public static Boolean convertToOCR()
        {
            //LicenseKey.LoadLicenseFile(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\eula\itextkey1629181689135_0.xml");
            var license_file = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\eula\itextkey1629181689135_0.xml";
            var pdfRenderPath = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\pdfrender-cli-1.0.1-exe-archive\pdfRender.exe";
            var nonOcrPdf = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\PDF_Input\nonocrresume.pdf";
            var imageOutput = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\Image_Output";

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // Instantiates the Windows Command Shell.
            startInfo.FileName = "cmd.exe";

            // The command line arguments for pdfRender.
            //    Type modifies the output file type (Important).
            startInfo.Arguments = "/c " + pdfRenderPath +
            " --pdf " + nonOcrPdf +
            " --out-dir " + imageOutput +
            " --type tiff" +
            " --scaling 2.0" +
            " --license " + license_file;

            startInfo.RedirectStandardOutput = true;
            startInfo.UseShellExecute = false;
            process.StartInfo = startInfo;
            process.Start();

            var output = process.StandardOutput.ReadToEnd();

            // Notifies via Command Line whether the PDF rendered successfully.
            Console.WriteLine(output);
            process.WaitForExit();

            IList<FileInfo> LIST_IMAGES_OCR = new List<FileInfo>();

            
            /*LIST_IMAGES_OCR.Add(new FileInfo(@"../sampleScannedImage.png"));
            LIST_IMAGES_OCR.Add(new FileInfo(@"../1.png"));
            LIST_IMAGES_OCR.Add(new FileInfo(@"../2.png"));
            LIST_IMAGES_OCR.Add(new FileInfo(@"../3.png"));*/

            var tesseractReader = new Tesseract4LibOcrEngine(tesseract4OcrEngineProperties);
            tesseract4OcrEngineProperties.SetPathToTessData(new FileInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\TESS_DATA"));

            var properties = new OcrPdfCreatorProperties();

            properties.SetPdfLang("en");
            var ocrPdfCreator = new OcrPdfCreator(tesseractReader, properties);
            var writer = new PdfWriter(new FileInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\PDF_Output\NewPdf.pdf"));

            ocrPdfCreator.CreatePdf(LIST_IMAGES_OCR, writer).Close();
            return true;
        }
        public static Boolean ReadWordFile(ClientContext cxt)
        {
            List mdl = cxt.Web.Lists.GetByTitle("Documents");
            CamlQuery q = new CamlQuery();
            string docname = "Auction Changes.docx";
            q.ViewXml = "<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='Text'>" + docname + "</Value></Eq></Where></Query></View>";

            ListItemCollection allfile = mdl.GetItems(q);
            cxt.Load(mdl);
            cxt.Load(allfile);
            cxt.ExecuteQuery();

            Boolean codeExist = false;

            foreach (Microsoft.SharePoint.Client.ListItem pdf in allfile)
            {
                try
                {


                    string fileName = cxt.Web.ServerRelativeUrl + "/Shared%20Documents/" + Convert.ToString(pdf.FieldValues["FileLeafRef"]);
                    var file = mdl.RootFolder.Files.GetByUrl(fileName);
                    cxt.Load(file);
                    cxt.ExecuteQuery();

                    ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                    cxt.ExecuteQuery();
                    byte[] sourceFileBytes = null;
                    System.IO.MemoryStream mStream = null;
                    mStream = new System.IO.MemoryStream();
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        sourceFileBytes = mStream.ToArray();
                        string b64String = Convert.ToBase64String(sourceFileBytes);
                    }

                    WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(mStream, false);
                    string wordcontent = wordprocessingDocument.MainDocumentPart.Document.Body.InnerText;
                    if (wordcontent.Contains("Auction pool"))
                    {
                        codeExist = true;
                        break;
                    }

                }
                catch
                {

                }
            }
            return codeExist;
        }

    }
}
