using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Document.NET;
using Xceed.Words.NET;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

namespace ParserTry1
{
    class Program
    {
        static void Main(string[] args)
        {


            Stopwatch stopwatch = new Stopwatch();
            //string Document1 = "start";

            List<string> Izveshenie = new List<string>();
            string path = @"C:\Users\Наталья\source\repos\ParserTry1\documents\";
            //string filename = @"C:\Users\Наталья\source\repos\ParserTry1\documents\1.doc";          
            string filename = "1.doc";

            Application app = new Application();
            app.Visible = false;

            //Convert(filename, path);
            //Console.WriteLine("Done");

            Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path + filename);
            try
            {
                Convert2txt(doc);
            }
            catch(Exception e) { Console.WriteLine(e.Message);
            }
            




            //Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path+filename);
            //stopwatch.Start();
            //Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path + filename);
            //stopwatch.Stop();
            Console.WriteLine("Open doc: " + stopwatch.ElapsedMilliseconds);

            Console.WriteLine(doc.Paragraphs.Count);
            try
            {
                for (int i = 1; i < doc.Paragraphs.Count; i++)
                {
                    stopwatch.Start();
                    Izveshenie.Add("\r\n" + doc.Paragraphs[i + 1].Range.Text);

                    stopwatch.Stop();

                    Console.WriteLine("List add: " + stopwatch.ElapsedMilliseconds);
                }

                doc.Close();

                foreach (string s in Izveshenie)
                {
                    Console.WriteLine(s);
                }
                // Console.WriteLine(text);
                Console.WriteLine("Done");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            //LoadXceed();
            Console.ReadLine();


            app.Quit();

        }
        //public static void Convert(string filename, string path)
        //{
        //    Application word = new Application();

        //    string fullpath = (path + filename);

        //    var sourceFile = new FileInfo(fullpath);
        //    Microsoft.Office.Interop.Word.Document document = word.Documents.OpenNoRepairDialog(sourceFile.FullName);
        //    string newFileName = sourceFile.FullName.Replace(".doc", ".docx");
        //    //string newFileName = $"{path}" + "new.docx";
        //    document.SaveAs2(newFileName, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: WdCompatibilityMode.wdWord2010);
        //    //document.Convert();
        //    word.ActiveDocument.Close();
        //    word.Quit();

        //}
        //public static void Convert2XML(string filename, string path)
        //{
        //    Application word = new Application();

        //    string fullpath = (path + filename);

        //    var sourceFile = new FileInfo(fullpath);
        //    Microsoft.Office.Interop.Word.Document document = word.Documents.OpenNoRepairDialog(sourceFile.FullName);
        //    string newFileName = sourceFile.FullName.Replace(".doc", ".xml");
        //    //string newFileName = $"{path}" + "new.docx";
        //    document.SaveAs2(sourceFile, WdSaveFormat.wdFormatXML);
        //    //document.Convert();
        //    word.ActiveDocument.Close();
        //    word.Quit();

        //}
        //public void LoadXceed()
        //{
        //    Xceed.Document.NET.Document doc = Xceed.Words.NET.DocX.Load("H:\\Винда март 2021\\repos\\Диплом\\new.docx");
        //    string txt = doc.Text;
        //    DateTime.Parse(txt);
        //}

        public static void Convert2txt(Microsoft.Office.Interop.Word.Document doc)
        {
            Application word = new Application();

            //string fullpath = (path + filename);

            var sourceFile = new FileInfo(doc.Path);
            Microsoft.Office.Interop.Word.Document document = doc;
            string newFileName = doc.FullName.Replace(".doc", ".txt");
            //string newFileName = $"{path}" + "new.docx";
            document.SaveAs2(newFileName, WdSaveFormat.wdFormatText);
            //document.Convert();
            
        }
    }
}
