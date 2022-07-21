using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using System.Reflection;
using System.IO;
using System.Runtime.InteropServices;

namespace COScan
{
    public class Program
    {
        #region console view / hide
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        #endregion
        
     
        string inputstring;
        string folderName;
        string NewFolderLoc;

        static void Main(string[] pdfilepath)
        {
            string sampletextpath = null;
            try
            {
                sampletextpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\PDFCompressor\\Sample.txt");
                Console.WriteLine("Staging path to Sample.txt............");

                Console.WriteLine("Path Found ");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Missing path to \\PDFCompressor\\Sample.txt !, Please create directory Accordingly.");

            }

            var handle = GetConsoleWindow();
            // Show
            ShowWindow(handle, SW_SHOW);
            //initialize paths
            #region 
            #endregion

            ///Call Functions to save path
            List<string> stringPath = new List<string>();
            

            stringPath.Add(sampletextpath);
            pdfilepath = stringPath.ToArray();

            //SaveToTextFile(pdfilepath, sampletextpath);

            //Call Funciton to Compress Files
            Console.WriteLine("Starting File Processing..............");
            ProcessFiles(pdfilepath, sampletextpath);

        }

        private static void ProcessFiles(string[] pdfilepath, string sampletextpath)
        {
            

            string savePath;
            //get directory with all scanned pdf files
            foreach (var t in pdfilepath)
            {
                Console.WriteLine("Reading path from Sample.txt..............");
                string[] pdfapthsstored = System.IO.File.ReadAllLines(t);
                
                foreach (var pdfdynamicfilepath in pdfapthsstored)
                {
                    var filename = pdfdynamicfilepath;
                    var p = Path.GetDirectoryName(pdfdynamicfilepath);
                savePath = p;


                string inputstring;
                string folderName = "";
                
                //get directory with all scanned pdf files
                List<string> pdffiles = new List<string>();

                pdffiles = Directory.GetFiles(savePath).Select(d => Path.GetFileName(d)).ToList();
                if (pdffiles.Count != 0)
                {

                    foreach (var r in pdffiles)
                    {
                        Console.WriteLine("Reading Path for File To File ....... :)");
                        //foreach file in directory confirm is all files are pdfs
                        string ext = Path.GetExtension(r);
                        if (ext == ".pdf" || ext == ".PDF")
                        {

                            inputstring = r;
                            int fileExtPos = inputstring.LastIndexOf(".");
                            if (fileExtPos >= 0)
                                inputstring = inputstring.Substring(0, fileExtPos);

                            Console.WriteLine("Loading Document for compression ....... :)");
                            //Load a existing PDF document
                            PdfLoadedDocument ldoc = new PdfLoadedDocument(filename);
                            //Create a new PDF compression options
                            Console.WriteLine("Setting Document Compression options ....... :)");

                            PdfCompressionOptions options = new PdfCompressionOptions();


                            Console.WriteLine("Setting image compression quality [Default is 15]..............");
                            options.CompressImages = true;
                            options.ImageQuality = 15;

                            Console.WriteLine("Optimising Document Font..............");
                            options.OptimizeFont = true;

                            Console.WriteLine("Optimising Page Content..............");
                            options.OptimizePageContents = true;

                            Console.WriteLine("Eliminating page MetaData..............");
                            options.RemoveMetadata = true;
                            Console.WriteLine("Compression options Loading completed");
                            Console.WriteLine("PassingParameters to Loaded document");
                            ldoc.CompressionOptions = options;

                            ldoc.FileStructure.IncrementalUpdate = false;
                            Console.WriteLine("Saving compressed file, please wait ..................");
                            MemoryStream ms = new MemoryStream();
                            ldoc.Save(ms);


                            // If directory does not exist, create it
                            if (!Directory.Exists(savePath))
                            {
                                Directory.CreateDirectory(savePath);
                            }
                            inputstring = savePath + "\\" + inputstring + ".pdf";
                            Console.WriteLine("Writing to Compressed file ..................");
                            File.WriteAllBytes(inputstring, ms.ToArray());


                            //update file name to start running

                            //ms.Dispose();
                            //Launching the PDF file using the default Application.[Acrobat Reader]
                            //System.Diagnostics.Process.Start(inputstring);
                            Console.WriteLine("Compresison completed :)");
                            //Console.WriteLine("Opening Compressed File for Preview ....... :)");


                            

                        }
                        else
                        {
                            Console.WriteLine("Hi User, Please note that no valid PDF files have been found.");
                        }
                    }

                }
            }
            }
            Console.WriteLine("File Compression Completed Successfully.");
            Console.WriteLine("Deleting Document logged path from Sample.txt");
            DeletePathCompleted(sampletextpath);
            Console.WriteLine("Done, Joel 2:12");
        }

        private static void DeletePathCompleted(string sampletextpath)
        {
            //string textfilepathsvar = @"C:\Users\ewaty\Desktop\PDFCompressor\Sample.txt";
            if (new FileInfo(sampletextpath).Length > 0)
            {
                // empty
                //File.AppendAllText(textfilepathsvar, Environment.NewLine + text);
                File.WriteAllText(sampletextpath, string.Empty);
            }
        }

        private static void SaveToTextFile(string[] pdfilepath, string sampletextpath)
        {
           
            foreach (var t in pdfilepath)
            {
                //checking if file exists in file
                if (System.IO.File.Exists(sampletextpath))
                {
                    var filename = Path.GetFileName(sampletextpath);
                    //writting content to file
                    WriteToFile(t);
                    //reading content to file
                    

                }
                else
                {
                    Console.WriteLine("File path Not Found");
                }

            }
        }

        private static void WriteToFile(string t)
        {
            string textfilepathsvar = @"C:\Users\ewaty\Desktop\PDFCompressor\Sample.txt";
            string text = t;
            if (new FileInfo(textfilepathsvar).Length > 0)
            {
                // empty
                File.AppendAllText(textfilepathsvar, Environment.NewLine + text);
            }
            if (new FileInfo(textfilepathsvar).Length == 0)
            {
                // empty
                File.WriteAllText(textfilepathsvar,text);
            }
        }
    }
}
