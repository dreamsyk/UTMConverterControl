using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;

namespace Maye
{
    public partial class ProcessBarX : DevComponents.DotNetBar.Metro.MetroForm
    {
        public ProcessBarX()
        {
            InitializeComponent();
        }

        private void ProcessBar_Load(object sender, EventArgs e)
        {
            this.progressBarX1.Minimum = 0;
            this.progressBarX1.Maximum = 100;
            this.progressBarX1.Value = 0;
            this.progressBarX1.Step = 1;
         }

        public void SetValue(int precent)
        {
            this.progressBarX1.Value = precent;
            if (precent == 100)
                this.Close();
        }

 

        private void ProcessBarX_FormClosing(object sender, FormClosingEventArgs e)
        {
            
        }
    }
}