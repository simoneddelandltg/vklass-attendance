using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VKlassAbsence
{
    public struct AbsenceProgress
    {
        public int TotalStudents { get; set; } = 0;
        public int FinishedStudents { get; set; } = 0;

        public string PathToOverview { get; set; } = "";

        public AbsenceProgress()
        {

        }
    }
}
