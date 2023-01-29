using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VKlassGrafiskFrånvaro
{
    class StringCaptureWriter : TextWriter
    {
        public override Encoding Encoding => throw new NotImplementedException();

        public override void WriteLine(string? value)
        {
            base.WriteLine(value);
            StringWrittenTo.Invoke(this, value);
        }

        public event EventHandler<string> StringWrittenTo;
    }


}
