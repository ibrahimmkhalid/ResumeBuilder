using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ResumeBuilder
{
    public partial class ResumeForm : Form
    {
        public string resumeType;
        public bool showUS;

        private string defaultResumeType;
        public ResumeForm()
        {
            InitializeComponent();
            this.resumeType = "";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.defaultResumeType = comboBox1.Text;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string resumeType = comboBox1.Text;
            if (resumeType == this.defaultResumeType) return;

            this.resumeType = resumeType;
            this.showUS = checkBox1.Checked;
            this.Close();
        }

        public async Task WaitForUser()
        {
            while (this.Visible)
            {
                await Task.Delay(100);
                Application.DoEvents();
            }
        }
    }
}
