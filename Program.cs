using Microsoft.Office.Interop.Word;
using ParserTry1;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
                switch (z.teacher.position)
                {
                    case "доц.":
                        z.teacher.position = "Доцент";
                        break;
                    case "ст.пр.":
                        z.teacher.position = "Старший преподаватель";
                        break;
                    case "асс.":
                        z.teacher.position = "Ассистент";
                        break;
                    case "проф.":
                        z.teacher.position = "Профессор";
                        break;
                }

                var teacherfullnameLower = z.teacher.fullname.ToLower();
               
                TextInfo myTI = new CultureInfo("ru-RU", false).TextInfo;
               z.teacher.fullname = myTI.ToTitleCase(teacherfullnameLower);
                Console.WriteLine(z.teacher.fullname);

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

           
            var jsonString = File.ReadAllText(writePath);
            //var notificationsFromJson = JsonSerializer.Deserialize<List<Notification>>(jsonString);

            app.Visible = true;
            var teacherCount = notifications.Count();
            
            Document docTable = app.Documents.Add();
            docTable.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            docTable.Paragraphs.Add();

            Range rng = docTable.Paragraphs[1].Range;

            rng.InsertBefore("РАСПИСАНИЕ ЗАНЯТИЙ ПРЕПОДАВАТЕЛЕЙ КАФЕДРЫ " +
                "ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ И ЗАЩИТЫ ИНФОРМАЦИИ " +
                "НА 1 - е ПОЛУГОДИЕ 2020 / 2021 УЧЕБНОГО ГОДА");
            rng.InsertParagraphAfter();
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 16;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
        
            rng.Tables.Add(docTable.Paragraphs[3].Range, teacherCount, 7, WdDefaultTableBehavior.wdWord9TableBehavior, WdAutoFitBehavior.wdAutoFitContent);

            Table tbl = docTable.Tables[1];
            tbl.Range.Font.Size = 10;
            
            tbl.Columns.DistributeWidth();
            tbl.Rows[1].Range.Font.Bold = 1;
            


            tbl.Cell(1, 1).Range.Text = "Ф.И.О. преподавателя";
            tbl.Cell(1, 2).Range.Text = "Понедельник";
            tbl.Cell(1, 3).Range.Text = "Вторник";
            tbl.Cell(1, 4).Range.Text = "Среда";
            tbl.Cell(1, 5).Range.Text = "Четверг";
            tbl.Cell(1, 6).Range.Text = "Пятница";
            tbl.Cell(1, 7).Range.Text = "Суббота";

            for(i = 2; i < teacherCount;)
            {
                foreach (var z in notifications)
                {
                    tbl.Cell(i, 1).Range.Text = $"{z.teacher.position} {z.teacher.fullname}";
                    i++;
                }

            }

            //tbl.Cell(2, 1).Range.Text = " ";
            //tbl.Cell(2, 2).Range.Text = " w ";

            //tbl.Cell(3, 1).Range.Text = "Author";
            //tbl.Cell(3, 2).Range.Text = " ww ";

            //using (StreamReader sr = new StreamReader(writePath))
            //{
            //    List<Notification> notifDes = JsonConvert.DeserializeObject<List<Notification>>(json);
            //    Console.WriteLine(notifDes);
            //}

            Console.ReadKey();
        }

        public static void Convert2txt(Document doc)
        {
            Application word = new Application();
            var sourceFile = new FileInfo(doc.Path);
            string newFileName = doc.FullName.Replace(".doc", ".txt");         
            doc.SaveAs2(newFileName, WdSaveFormat.wdFormatText);
        }
    }
}

