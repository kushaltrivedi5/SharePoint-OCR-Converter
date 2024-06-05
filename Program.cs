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
        
        static void Main(string[] args)
        {
            try
            {
                var codefound = convertToOCR();
                if (codefound == true)
                {
                    //PDF APPROVED CODE GOES HERE
                }
                Console.WriteLine(codefound);
                Console.Read();

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
       public static Boolean convertToOCR()
        {
            Tesseract4OcrEngineProperties tesseract4OcrEngineProperties = new Tesseract4OcrEngineProperties();
            
            //LicenseKey.LoadLicenseFile(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\eula\itextkey1629181689135_0.xml");
            var license_file = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\eula\itextkey1629181689135_0.xml";
            var pdfRenderPath = @"C:\Users\kushal.trivedi.vc\Desktop\Projects\pdfRender\pdfrender-cli-1.0.1-exe-archive\pdfRender.exe";
            DirectoryInfo nonOcrPdf = new DirectoryInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\PDF_Input\nonocrresume.pdf");
            DirectoryInfo imgDirectory = new DirectoryInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\Image_Output");

            for (var i = 0; i < imgDirectory.GetFiles().Length; i++)
            {
                imgDirectory.GetFiles()[i].Delete();
            }

            Process process = new Process();
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.WindowStyle = ProcessWindowStyle.Hidden;

            // Instantiates the Windows Command Shell.
            startInfo.FileName = "cmd.exe";

            // The command line arguments for pdfRender.
            //    Type modifies the output file type (Important).
            startInfo.Arguments = "/c " + pdfRenderPath +
            " --pdf " + nonOcrPdf.FullName +
            " --out-dir " + imgDirectory.FullName +
            " --type tiff" +
            " --scaling 3.0" +
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

            for (var i = 0; i < imgDirectory.GetFiles().Length; i++)
            {
                LIST_IMAGES_OCR.Add(imgDirectory.GetFiles()[i]); 
            }

            var tesseractReader = new Tesseract4LibOcrEngine(tesseract4OcrEngineProperties);
            tesseract4OcrEngineProperties.SetPathToTessData(new FileInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\TESS_DATA"));

            var properties = new OcrPdfCreatorProperties();

            properties.SetPdfLang("en");
            var ocrPdfCreator = new OcrPdfCreator(tesseractReader, properties);
            var writer = new PdfWriter(new FileInfo(@"C:\Users\kushal.trivedi.vc\Desktop\Projects\PDFtoOCR\bin\PDF_Output\NewPdf.pdf"));

            ocrPdfCreator.CreatePdf(LIST_IMAGES_OCR, writer).Close();
            return true;
        }
    }
}
