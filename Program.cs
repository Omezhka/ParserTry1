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

            string path = AppDomain.CurrentDomain.BaseDirectory + @"documents\";
            string pathOutput = AppDomain.CurrentDomain.BaseDirectory + @"outputDocuments\";
            //путь для выходных
            string filename = path + "1.doc";

            string filenametxt = pathOutput + "1.txt";

            Application app = new Application();
            app.Visible = false;

            Document doc = app.Documents.OpenNoRepairDialog(filename);
            try
            {
                Convert2txt(doc);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            app.ActiveDocument.Close();

            using (StreamReader sr = new StreamReader(filenametxt, System.Text.Encoding.Default))
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
            //ну тут меняю сокращенные позишны на полные, шоб в расписании красиво выглядело
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

            string writePath = pathOutput + "hta.txt";

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

            //var jsonString = File.ReadAllText(writePath);
            //var notificationsFromJson = JsonSerializer.Deserialize<List<Teacher>>(jsonString, options1);


            app.Visible = true;

            var teacherCount = notifications.Count();
            //тут создаю новый док, задаю ему альбомную ориентацию
            Document docTable = app.Documents.Add();
            docTable.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            // тут какие-то замуты с параграфами, я не до конца выкупила, но если строчку ниже убрать,
            // то таблица принимает стили текста перед таблицей 
            docTable.Paragraphs.Add();

            Range rng = docTable.Paragraphs[1].Range;

            rng.InsertBefore("РАСПИСАНИЕ ЗАНЯТИЙ ПРЕПОДАВАТЕЛЕЙ КАФЕДРЫ " +
                "ИНФОРМАЦИОННЫХ ТЕХНОЛОГИЙ И ЗАЩИТЫ ИНФОРМАЦИИ " +
                "НА 1 - е ПОЛУГОДИЕ 2020 / 2021 УЧЕБНОГО ГОДА");
            rng.InsertParagraphAfter();
            //собсна, стили текста выше
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 16;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
            //добавление таблицы в третий параграф, кол-во строк = кол-во нотификейшнов, 7 столбцов, дальше какие-то настройки таблицы
            rng.Tables.Add(docTable.Paragraphs[3].Range, teacherCount + 1, 7, WdDefaultTableBehavior.wdWord9TableBehavior, WdAutoFitBehavior.wdAutoFitWindow);

            //тут всё закручено на Range
            Table tbl = docTable.Tables[1];
            tbl.Range.Font.Size = 10;
            //это настройки таблицы
            tbl.Columns.DistributeWidth();
            tbl.Rows[1].Range.Font.Bold = 1;
            // ну, заполнение шапки таблицы))
            tbl.Cell(1, 1).Range.Text = "Ф.И.О. преподавателя";
            tbl.Cell(1, 2).Range.Text = "Понедельник";
            tbl.Cell(1, 3).Range.Text = "Вторник";
            tbl.Cell(1, 4).Range.Text = "Среда";
            tbl.Cell(1, 5).Range.Text = "Четверг";
            tbl.Cell(1, 6).Range.Text = "Пятница";
            tbl.Cell(1, 7).Range.Text = "Суббота";
            // этот список для раскидывания занятий по ячейкам соотв. дня недели
            var week = new List<string>
            {
                "пнд",
                "втp",
                "сpд",
                "чтв",
                "птн",
                "сбт"
            };
            // вывод преподов 
            for (i = 2; i <= teacherCount + 1;)
            {
                tbl.Cell(i, 1).Range.Text = $"{notifications[i - 2].teacher.position} {notifications[i - 2].teacher.fullname}";
                i++;

            }
            //tbl.Cell(2,2).Split(2);
            // тут как раз таки заполнение таблицы, тут нафигачила ненужных списков(как мне кажется) пар по четной и нечетной неделе
            // пытаюсь вывести расписание по четности недели
            for (i = 0; i < teacherCount; i++) //столбец
            {
                for (var k = 0; k < notifications[i].scheduleList.Count; k++) // строка
                {
                    int indexDayPosition = 0;
                    var trueWeekSchedule = new List<string>();
                    var falseWeekSchedule = new List<string>();
                    indexDayPosition = week.IndexOf(notifications[i].scheduleList[k].days);

                    if (notifications[i].scheduleList[k].Week)
                    {
                        trueWeekSchedule.Add($"{notifications[i].scheduleList[k].classhours}" +
                                             $" {notifications[i].scheduleList[k].group}" +
                                             $" {"a." + notifications[i].scheduleList[k].audience}");

                        //    tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter("чет: " + $"{notifications[i].scheduleList[k].classhours}" +
                        //                                                                     $" {notifications[i].scheduleList[k].group}" +
                        //                                                                     $" {"a." + notifications[i].scheduleList[k].audience}\r\n");
                    }
                    else
                    {
                        falseWeekSchedule.Add($"{notifications[i].scheduleList[k].classhours}" +
                                             $" {notifications[i].scheduleList[k].group}" +
                                             $" {"a." + notifications[i].scheduleList[k].audience}");

                        //tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter("нечет: " + $"{notifications[i].scheduleList[k].classhours} " +
                        //                                                                   $"{notifications[i].scheduleList[k].group} " +
                        //                                                                   $"{"a." + notifications[i].scheduleList[k].audience}\r\n");
                    }
                    // сейчас выводит пары по четной недели, потому что у меня была такая тактика, и я ее придерживалась
                    foreach (var w in trueWeekSchedule)
                    {
                        tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter("zhopa");

                        tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter($"{notifications[i].scheduleList[k].classhours}" +
                                                                                         $" {notifications[i].scheduleList[k].group}" +
                                                                                        $" {"a." + notifications[i].scheduleList[k].audience}\r\n");
                    }

                }
            }




            //app.Visible = false;
            // тут я создаю новый документ с расписанием одного препода
            // не пониманию как их создать много

            Document teacherScheduleTable = app.Documents.Add();
            //тут как раз таки плодится много документов с фио каждого препода
            var nameOfFile = new List<string>();


            foreach (var t in notifications)
            {
                var filenamedd = pathOutput + @"prepods\" + t.teacher.fullname + ".docx";
                teacherScheduleTable.SaveAs2(filenamedd, WdSaveFormat.wdFormatDocumentDefault);
                nameOfFile.Add(System.IO.Path.GetFileName(filenamedd));

            }
            // а рисуется только в последнем созданном доке
            // свойства
            teacherScheduleTable.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            teacherScheduleTable.Paragraphs.Add();
            //это по аналогии с предыдущим доком
            var classhours = new List<string>
                {
                    "08.30-10.00",
                    "10.10-11.40",
                    "11.50-13.20",
                    "13.50-15.20",
                    "15.30-17.00",
                    "17.10-18.40"
                };

            Range rngtst = teacherScheduleTable.Paragraphs[1].Range;

            rngtst.InsertBefore("РАСПИСАНИЕ ВАШИХ ЗАНЯТИЙ НА 1 - е ПОЛУГОДИЕ 2020 / 2021 УЧЕБНОГО ГОДА");
            rngtst.InsertParagraphAfter();
            rngtst.Font.Name = "Times New Roman";
            rngtst.Font.Size = 16;
            rngtst.Font.Bold = 1;
            rngtst.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;



            rngtst.Tables.Add(teacherScheduleTable.Paragraphs[3].Range, classhours.Count + 1, 7, WdDefaultTableBehavior.wdWord9TableBehavior, WdAutoFitBehavior.wdAutoFitWindow);

            Table tbltst = teacherScheduleTable.Tables[1];
            tbltst.Range.Font.Size = 10;

            tbltst.Columns.DistributeWidth();

            tbltst.Rows[1].Range.Font.Bold = 1;

            tbltst.Cell(1, 1).Range.Text = " ";
            tbltst.Cell(1, 2).Range.Text = "Понедельник";
            tbltst.Cell(1, 3).Range.Text = "Вторник";
            tbltst.Cell(1, 4).Range.Text = "Среда";
            tbltst.Cell(1, 5).Range.Text = "Четверг";
            tbltst.Cell(1, 6).Range.Text = "Пятница";
            tbltst.Cell(1, 7).Range.Text = "Суббота";


            for (i = 2; i <= classhours.Count + 1;)
            {
                tbltst.Cell(i, 1).Range.Text = classhours[i - 2];
                i++;

            }
            //tbl.Cell(2,2).Split(2);
            //тут та же проблема с определением четности
            for (i = 0; i < classhours.Count; i++) //столбец
            {
                for (var k = 0; k < notifications[i].scheduleList.Count; k++) // строка
                {
                    var indexDayPosition = 0;
                    var indexClasshoursPosition = 0;
                    var trueWeekSchedule = new List<string>();
                    var falseWeekSchedule = new List<string>();
                    indexDayPosition = week.IndexOf(notifications[i].scheduleList[k].days);
                    indexClasshoursPosition = classhours.IndexOf(notifications[i].scheduleList[k].classhours);

                    if (notifications[i].scheduleList[k].Week)
                    {
                        trueWeekSchedule.Add($" { notifications[i].teacher.fullname} {notifications[i].scheduleList[k].group}" +
                                             $" {"a." + notifications[i].scheduleList[k].audience}");

                        //    tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter("чет: " + $"{notifications[i].scheduleList[k].classhours}" +
                        //                                                                     $" {notifications[i].scheduleList[k].group}" +
                        //                                                                     $" {"a." + notifications[i].scheduleList[k].audience}\r\n");
                    }
                    else
                    {
                        falseWeekSchedule.Add($" {notifications[i].scheduleList[k].group}" +
                                             $" {"a." + notifications[i].scheduleList[k].audience}");

                                //tbl.Cell(i + 2, indexDayPosition + 2).Range.InsertAfter("нечет: " + $"{notifications[i].scheduleList[k].classhours} " +
                                //                                                                   $"{notifications[i].scheduleList[k].group} " +
                                //                                                                   $"{"a." + notifications[i].scheduleList[k].audience}\r\n");
                            }
                            foreach (var w in trueWeekSchedule)
                            {
                                foreach (var item in nameOfFile)
                                {
                                    if ((notifications[i].teacher.fullname + ".docx").Equals(item[i]))
                                        tbltst.Cell(indexClasshoursPosition + 2, indexDayPosition + 2).Range.InsertAfter($" { notifications[i].teacher.fullname} {notifications[i].scheduleList[k].group}" +
                                                                                                            $" {"a." + notifications[i].scheduleList[k].audience}\r\n");
                                }
                                    //tbltst.Cell(i+1, indexDayPosition+2).Range.InsertAfter("zhopa");
                                
                            }
                        }
                    

            }


            Console.ReadKey();
        }

        public static void Convert2txt(Document doc)
        {
            string newFileName = doc.FullName.Replace(".doc", ".txt");
            doc.SaveAs2(newFileName, WdSaveFormat.wdFormatText);
        }
    }
}

