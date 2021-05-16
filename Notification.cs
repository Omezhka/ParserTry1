using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ParserTry1
{
    /// <summary>
    /// собственно извещение, у него свойство препод , которое берётся из заголовка
    /// </summary>
    public class Notification
    {
        public Notification(List<string> izvItem)
        {
            int i = 0;

            // добавляем преподавателя из заголовка извещения

            var regTeacher = new Regex(Pattern.header);
            var regSchdlItem = new Regex(Pattern.scheduleItem);

            teacher = new Teacher()
            {
                position = regTeacher.Match(izvItem[i]).Groups["position"].ToString(),
                fullname = regTeacher.Match(izvItem[i]).Groups["fullname"].ToString(),
                cathedra = regTeacher.Match(izvItem[i]).Groups["cathedra"].ToString(),
            };

            //наполняем расписание

            scheduleList = new List<Schedule>();

            i = 9; // пропускаем неинтересные строчки

            while (i < izvItem.Count)
            {
                if (Regex.IsMatch(izvItem[i], Pattern.scheduleItem))
                {
                    scheduleList.Add(new Schedule(izvItem[i])
                    {
                        Week = false
                    });
                    i++;
                }
                else
                {
                    i += 3; // если нечетная недля пустая, то пропускаем 3 строки и попадаем на чётную, следующий цикл
                    break;
                }
            }
            while (i < izvItem.Count)
            {
                if (Regex.IsMatch(izvItem[i], Pattern.scheduleItem))
                {
                    scheduleList.Add(new Schedule(izvItem[i])
                    {
                        Week = true
                    });
                    i++;
                }
                else
                {
                    break;
                }
            }

        }
        /// <summary>
        /// препод
        /// </summary>
        public Teacher teacher { get; set; }
        /// <summary>
        /// список строк из извещения
        /// </summary>
        public List<Schedule> scheduleList { get; set; }
    }
}
