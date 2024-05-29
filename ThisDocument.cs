using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Word;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using ResumeBuilder.Properties;

namespace ResumeBuilder
{
    public partial class ThisDocument
    {
        private async void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            float inchesPointFive = Application.InchesToPoints(0.5F);
            this.PageSetup.BottomMargin = inchesPointFive;
            this.PageSetup.TopMargin = inchesPointFive;
            this.PageSetup.RightMargin = inchesPointFive;
            this.PageSetup.LeftMargin = inchesPointFive;

            ResumeForm FormData = new ResumeForm();
            FormData.Show();
            await FormData.WaitForUser();
            string resumeType = FormData.resumeType;
            bool ShowUS = FormData.showUS;

            //Heading
            object start = 0;
            object end = 0;
            Word.Range rng = this.Range(ref start, ref end);
            rng.Text = Resources.Name;
            rng.Font.Size = 20;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            start = end = rng.End;
            rng = this.Range(ref start, ref end);
            if (ShowUS)
            {
                rng.Text = " (US Citizen)\n";
                rng.Font.Size = 7;
                rng.Font.Bold = 0;
                rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            } else
            {
                rng.Text = "\n";
            }

            start = end = rng.End;
            rng = this.Range(ref start, ref end);
            rng.Text = Resources.subtitle + "\n";
            rng.Font.Size = 10;
            rng.Font.Bold = 0;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            rng = this.AddSection("Education", rng.End, resumeType);
            rng = this.AddSection("Experiences", rng.End, resumeType);
            rng = this.AddSection("Projects", rng.End, resumeType);

            //end
            foreach (Word.Paragraph p in this.Paragraphs)
            {
                p.Space1();
                p.SpaceAfter = 0;
                p.SpaceBefore = 0;
                p.Range.Font.Name = "Calibri";
            }
        }

        private Word.Range AddSection(string sectionName, int index, string type)
        {
            object start = index;
            object end = index;
            Word.Range rng = this.Range(ref start, ref end);
            rng.Text = sectionName + "\n";
            rng.Font.Size = 14;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
            return rng;
        }

        private void ThisDocument_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisDocument_Startup);
            this.Shutdown += new System.EventHandler(ThisDocument_Shutdown);
        }

        #endregion
    }
}
