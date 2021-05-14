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
using System.Text.RegularExpressions;

namespace ParserTry1
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            string PATTERN2 = @"\w*И\sЗ\sВ\sЕ\sЩ\sЕ\sН\sИ\sЕ\s*(?<prepod>.*\.)(?<kafedra>.*)";
            string PATTERN3 = @"\s+НЕЧЕТНАЯ\sНЕДЕЛЯ";
            string PATTERN4 = @"\s+ЧЕТНАЯ\sНЕДЕЛЯ";
            string PATTERN5 = @"\u00A6\s(?<subject>[\w,\W]*)\u00A6\s(?<days>\w\w\w)\s\u00A6(?<classhours>.*)\s*\u00A6(?<audience>.*)\s*\u00A6(?<group>.*)\s*\u00A6";

            Stopwatch stopwatch = new Stopwatch();
            //string Document1 = "start";

            List<string> Izveshenie = new List<string>();
            string path = @"C:\Users\Наталья\source\repos\ParserTry1\documents\";
            //string filename = @"C:\Users\Наталья\source\repos\ParserTry1\documents\1.doc";          
            string filename = "1.doc";
            string filenametxt = "1.txt"; 

            Application app = new Application();
            app.Visible = false;

            //Convert(filename, path);
            //Console.WriteLine("Done");

            Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path + filename);
            try
            {
                Convert2txt(doc);
            }
            catch(Exception e) { 
                Console.WriteLine(e.Message);
            }
            app.ActiveDocument.Close();

            using (StreamReader sr = new StreamReader(path+filenametxt, System.Text.Encoding.Default))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    Izveshenie.Add("\r\n" + line);
                    Console.WriteLine(line);
                }
            }
       
            var reg = new Regex(PATTERN2);
            var groupNames = reg.GetGroupNames();
            var reg2 = new Regex(PATTERN5);
            var groupNames2 = reg2.GetGroupNames();

            foreach (var s in Izveshenie)
            {
                if (reg.IsMatch(s))
                {
                    foreach (var t in groupNames)
                    {
                        Console.WriteLine($"{t} : [{reg.Match(s).Groups[t]}]");
                    }


                }
                if (reg2.IsMatch(s))
                {
                    foreach (var t in groupNames2)
                    {
                        Console.WriteLine($"{t} : [{reg2.Match(s).Groups[t].ToString().Trim()}]");
                    }


                }

            }


            //Microsoft.Office.Interop.Word.Document doctxt = app.Documents.OpenNoRepairDialog(path + filenametxt);


            //Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path+filename);
            //stopwatch.Start();
            //Microsoft.Office.Interop.Word.Document doc = app.Documents.OpenNoRepairDialog(path + filename);
            //stopwatch.Stop();
            //Console.WriteLine("Open doc: " + stopwatch.ElapsedMilliseconds);

            //Console.WriteLine(doctxt.Paragraphs.Count);
            //try
            //{
            //    for (int i = 1; i < doctxt.Paragraphs.Count; i++)
            //    {
            //        stopwatch.Start();
            //        Izveshenie.Add("\r\n" + doctxt.Paragraphs[i + 1].Range.Text);

            //        stopwatch.Stop();

            //        Console.WriteLine("List add: " + stopwatch.ElapsedMilliseconds);
            //    }

            //    doc.Close();

            //    foreach (string s in Izveshenie)
            //    {
            //        Console.WriteLine(s);
            //    }
            //    // Console.WriteLine(text);
            //    Console.WriteLine("Done");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine(ex.Message);
            //}

            ////LoadXceed();
            Console.ReadLine();




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
