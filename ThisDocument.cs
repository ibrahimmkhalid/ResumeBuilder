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
        private string resumeType;

        private struct sectionParse
        {
            public sectionParse(string[] s, List<string> p, bool i)
            {
                data = s;
                points = p;
                include = i;
            }

            public string[] data { get; set; }
            public List<string> points { get; set; }
            public bool include { get; set; }


        }

        private async void ThisDocument_Startup(object sender, System.EventArgs e)
        {
            float inchesPointFive = Application.InchesToPoints(0.5F);
            this.PageSetup.BottomMargin = inchesPointFive;
            this.PageSetup.TopMargin = inchesPointFive;
            this.PageSetup.RightMargin = inchesPointFive;
            this.PageSetup.LeftMargin = inchesPointFive;

            //ResumeForm FormData = new ResumeForm();
            //FormData.Show();
            //await FormData.WaitForUser();
            //this.resumeType = FormData.resumeType;
            //bool ShowUS = FormData.showUS;

            this.resumeType = "AI";
            bool ShowUS = true;

            //Heading
            object start = 0;
            object end = 0;
            Word.Range rng = this.Range(ref start, ref end);
            rng.Text = Resources.Name;
            rng.Font.Size = 20;
            rng.Font.Bold = 1;
            foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
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
            foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;

            start = end = rng.End;
            rng = this.Range(ref start, ref end);
            rng.Text = Resources.subtitle + "\n";
            foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
            rng.Font.Size = 10;
            rng.Font.Bold = 0;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;


            rng = this.AddSection("Education", rng.End);
            rng = this.AddSection("Experiences", rng.End);
            rng = this.AddSection("Projects", rng.End);

            //end
            foreach (Word.Paragraph p in this.Paragraphs)
            {
                p.Space1();
                p.Range.Font.Name = "Calibri";
            }
        }

        private Word.Range AddSection(string sectionName, int index)
        {
            object start = index;
            object end = index;
            Word.Range rng = this.Range(ref start, ref end);
            rng.Text = sectionName + "\n";
            foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
            foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = 8;
            rng.Font.Size = 14;
            rng.Font.Bold = 1;
            rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

            string dataCSV;
            if (sectionName == "Education")
            {
                dataCSV = Resources.education;
            } 
            else if (sectionName == "Experience")
            {
                dataCSV = "blekh";
            } 
            else
            {
                dataCSV = "blekh";
            }

            float pageWidth = this.Application.Selection.PageSetup.PageWidth;
            float leftMargin = this.Application.Selection.PageSetup.LeftMargin;
            float rightMargin = this.Application.Selection.PageSetup.RightMargin;
            float usableWidth = pageWidth - leftMargin - rightMargin;

            string[] entries = dataCSV.Split('~');
            foreach(string s in entries)
            {
                sectionParse section = ParseCSV(s);
                if (section.include)
                {
                    start = end = rng.End;
                    object linestart = start;
                    rng = this.Range(ref start, ref end);

                    rng.ParagraphFormat.TabStops.ClearAll();
                    object leftTabPosition = this.Application.InchesToPoints(1); // Adjust as needed
                    object rightTabPosition = this.Application.InchesToPoints(usableWidth / 72); // Adjust as needed
                    rng.ParagraphFormat.TabStops.Add(((float)leftTabPosition), Word.WdTabAlignment.wdAlignTabLeft, Word.WdTabLeader.wdTabLeaderSpaces);
                    rng.ParagraphFormat.TabStops.Add(((float)rightTabPosition), Word.WdTabAlignment.wdAlignTabRight, Word.WdTabLeader.wdTabLeaderSpaces);

                    string name = section.data[0] + (section.data[1] == "" ? "" : ", ");
                    string location =section.data[1] == "" ? "" : section.data[1];
                    string date = section.data[2];
                    string line = name + location + "\t" + date;
                    rng.Text = line;
                    rng.Font.Size = 10;

                    rng = this.Range(ref start, ref end);
                    rng.End = rng.Start + name.Length;
                    rng.Font.Bold = 1;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    start = rng.End;
                    end = (int)start + location.Length;
                    rng = this.Range(ref start, ref end);
                    rng.Font.Bold = 0;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    start = rng.End + 1;
                    end = (int)start + date.Length;
                    rng = this.Range(ref start, ref end);
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                    rng.Font.Bold = 0;

                    start = linestart;
                    end = (int)start + line.Length;
                    rng = this.Range(ref start, ref end);
                    rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                    rng.Text = rng.Text + "\n";
                    foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
                    foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = 4;

                    start = end = rng.End;
                    rng = this.Range(ref start, ref end);
                    rng.Text = section.data[3] == "" ? "" : section.data[3] + "\n";
                    foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
                    rng.Font.Bold = 0;
                    rng.Font.Size = 10;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;

                    start = end = rng.End;
                    rng = this.Range(ref start, ref end);
                    string points = "";
                    foreach(string point in section.points)
                    {
                        points = points + point + "\n";
                    }
                    rng.Text = points;
                    foreach (Word.Paragraph p in rng.Paragraphs) p.SpaceBefore = p.SpaceAfter = 0;
                    rng.Font.Bold = 0;
                    rng.Font.Size = 10;
                    rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                    rng.ListFormat.ApplyBulletDefault();

                  }
            }
            return rng;
        }

        private sectionParse ParseCSV(string s)
        {
            sectionParse sectionData = new sectionParse();
            sectionData.data = new string[4];
            sectionData.points = new List<string>();
            sectionData.include = false;

            if (s == "blekh") //TODO: REMOVE
            {
                return sectionData;
            }

            char[] trimChars = { '\r', '\n' , ' ', '\t', ',', '\"', '\'', '\0', '\\'};

            bool firstline = true;
            string[] source = s.
                Trim(trimChars)
                .Split('\n');
            string[] header = source[0]
                .Trim(trimChars)
                .Split(',');
            if (!header[3].Contains(this.resumeType)) return sectionData;
            sectionData.include = true;


            sectionData.data[0] = header[0].Trim(trimChars); //name
            sectionData.data[1] = header[1].Trim(trimChars); //location?
            sectionData.data[2] = header[2].Trim(trimChars); //date


            if (header.Length > 5)
            {
                for (int i = 4; i < header.Length; i++)
                {
                    sectionData.data[3] += header[i];
                }
            } 
            else if (header.Length < 5)
            {
                sectionData.data[3] = "";
            } 
            else
            {
                sectionData.data[3] = header[4];
            }
            sectionData.data[3] = sectionData.data[3].Trim(trimChars); // headline?

            foreach (string point in source)
            {
                if (firstline)
                {
                    firstline = false;
                    continue;
                }
                sectionData.points.Add(point.Trim(trimChars));
            }
            return sectionData;
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
