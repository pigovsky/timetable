using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParseTimetableFromExcel
{
    class CentralExceptionProcessor
    {
        private static StreamWriter log = File.CreateText("Errors.log");
        public static void process(object e)
        {
            log.WriteLine(e.ToString());
        }
    }
}
