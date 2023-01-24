using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VKlassAbsence
{
    internal class Lesson
    {
        public DateTime StartTime { get; set; }
        public DateTime StopTime { get; set; }
        public string Course { get; set; }
        public LessonStatus Status { get; set; }
        public int MissingMinutes { get; set; }
    }
}
