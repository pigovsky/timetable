using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ParseTimetableFromExcel
{
    public partial class ProgressForm : Form
    {
        private long _currentNumberOfIterationsPass;

        public long currentNumberOfIterationsPass 
        {
            get
            {
                return _currentNumberOfIterationsPass;
            }
            set
            {
                _currentNumberOfIterationsPass = value;
                UpdateProgress();
            }
        }
        public long totalNumberOfIterations { get; set; }


        public ProgressForm()
        {
            InitializeComponent();            
        }

        delegate void UpdateProgressCallback();
       
        private void UpdateProgress()
        {
            if (importProgressBar.InvokeRequired)
                Invoke(new UpdateProgressCallback(UpdateProgress));
            else
                importProgressBar.Value = (int)
                    (currentNumberOfIterationsPass*100/
                    totalNumberOfIterations);
        }

        delegate void HideImportProgressFormCallback(ThreadStart fun);

        public void HideProgressForm(ThreadStart fun)
        {
            // InvokeRequired required compares the thread ID of the 
            // calling thread to the thread ID of the creating thread. 
            // If these threads are different, it returns true. 
            if (this.InvokeRequired)
            {
                this.Invoke(new
                    HideImportProgressFormCallback(HideProgressForm));
            }
            else
            {
                fun();
                this.Hide();
            }
        }
    }
}
