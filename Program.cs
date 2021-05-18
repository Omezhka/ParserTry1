using Microsoft.Office.Interop.Word;
using ParserTry1;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text.Unicode;

namespace regexpParse
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            List<string> izv = new List<string>();
        
            string path = @"C:\Users\Наталья\source\repos\ParserTry1\documents\";
            //string filename = @"C:\Users\Наталья\source\repos\ParserTry1\documents\1.doc";          
            string filename = "1.doc";
            string filenametxt = "1.txt";

            Application app = new Application();
            app.Visible = false;

            Document doc = app.Documents.OpenNoRepairDialog(path + filename);
            try
            {
                Convert2txt(doc);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            app.ActiveDocument.Close();

            using (StreamReader sr = new StreamReader(path + filenametxt, System.Text.Encoding.Default))
            {
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    izv.Add(line);
                    
                }
            }
     
            var regHeader = new Regex(Pattern.header);
            var groupHeaderNames = regHeader.GetGroupNames();
            var regScheduleItem = new Regex(Pattern.scheduleItem);
            var groupScheduleItemNames = regScheduleItem.GetGroupNames();

            int i = 0;

            var notifications = new List<Notification>();

            while (i < izv.Count)
            { 
                if (regHeader.IsMatch(izv[i]))
                {
                    var izvHeaderCathedra = regHeader.Match(izv[i]).Groups["cathedra"].ToString(); // берём название кафедры из заголовка
                    var izvItem = new List<string>();
                    while (izv[i] != "         Специалист отдела ОУП и ККО Бусова О.В.")
                    {
                        izvItem.Add(izv[i]);
                        i++;
                    }
                    if (izvHeaderCathedra == "Информационных технологий и за") // сравниваем название кафедры с нужной, и если совпало - добавляем в список распарщеных извещений
                    {
                        notifications.Add(new Notification(izvItem));
                    }
                }
                i++;
            }

            foreach (var z in notifications)
            {            
                    Console.WriteLine($"{z.teacher.position} {z.teacher.fullname} {z.teacher.cathedra}");
                    foreach (var y in z.scheduleList)
                    {
                        Console.WriteLine($"{y.group} {y.Week}");
                    }
            }

            var options = new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.Create(UnicodeRanges.BasicLatin, UnicodeRanges.Cyrillic),
                WriteIndented = true
            };

            string json = JsonSerializer.Serialize(notifications, options);
            //Console.WriteLine(json);

            string writePath = @"C:\Users\Наталья\source\repos\ParserTry1\documents\hta.txt";

            try
            {
                using (StreamWriter sw = new StreamWriter(writePath, false, System.Text.Encoding.Default))
                {
                    sw.WriteLine(json);
                }

                Console.WriteLine("Запись выполнена");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            Console.ReadKey();
        }

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

