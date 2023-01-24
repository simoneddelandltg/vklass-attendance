using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VKlassAbsence
{
    internal class AttendanceData
    {
        public string Name { get; set; }
        public double Attendance { get; set; }
        public double ValidAbsence { get; set; }
        public double InvalidAbsence { get; set; }
        public List<Lesson> Lessons { get; set; } = new List<Lesson>();
    }
}
