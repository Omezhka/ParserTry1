using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ParserTry1
{
    public class Schedule
    {
        public Schedule() { }
        public Schedule(string s)
        {
            var schedulerRegex = new Regex(Pattern.scheduleItem);

            //Week = schedulerRegex.Match(s).Groups["week"];                      присвоить в другом месте
            subject = schedulerRegex.Match(s).Groups["subject"].ToString().Trim();
            days = schedulerRegex.Match(s).Groups["days"].ToString().Trim();
            classhours = schedulerRegex.Match(s).Groups["classhours"].ToString().Trim();
            audience = schedulerRegex.Match(s).Groups["audience"].ToString().Trim();
            group = schedulerRegex.Match(s).Groups["group"].ToString().Trim();
        }
        /// <summary>
        /// тип недели чёт/нечёт (true/false)
        /// </summary>        
        public bool Week { get; set; }
        public string subject { get; set; }
        public string days { get; set; }
        public string classhours { get; set; }
        public string audience { get; set; }
        public string group { get; set; }
    }
}
