using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParseTimetableFromExcel
{
    class CentralExceptionProcessor
    {
        public static void process(object e)
        {
            MessageBox.Show(e.ToString());
        }
    }
}
