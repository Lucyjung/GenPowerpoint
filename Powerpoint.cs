using System;
using System.Collections.Generic;
using System.IO;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Drawing;
using System.Text.RegularExpressions;

namespace Report1
{
    class Powerpoint
    {
        static Presentation pptPresentation;
        static Application pptApplication;
        static string[] dynamicField;
        public static void initialReport(string outputFile)
        {
            pptApplication = new Application();
            
            string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template.pptx");
            File.Copy(templatePath, outputFile, true);
            // Create the Presentation File
            pptPresentation = pptApplication.Presentations.Open(outputFile);
            dynamicField = Config.dynamicVar.Split(',');
        }
        public static void CloseReport(string outputFile, int slideIndex = 1)
        {
            pptPresentation.Slides[slideIndex].Delete();
            //pptPresentation.SaveCopyAs(outputFile, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            pptPresentation.Save();
            pptPresentation.Close();
            pptApplication.Quit();
            //Marshal.FinalReleaseComObject(pptApplication);
        }
        public static void genReport(string processName, string period, string year, System.Data.DataTable dt, int slideIndex = 1, bool isChart=false)
        {
            if (isChart)
            {
                replaceChartVariable(pptPresentation.Slides[slideIndex].Shapes, dt);
            } else
            {
                pptPresentation.Slides[slideIndex].Duplicate();
                pptPresentation.Slides[slideIndex].Select();
                replaceVariable(pptPresentation.Slides[slideIndex].Shapes, "PROCESSNAME", processName);
                replaceVariable(pptPresentation.Slides[slideIndex].Shapes, "PERIOD", period);
                replaceVariable(pptPresentation.Slides[slideIndex].Shapes, "YEAR", year);
                replaceVariable(pptPresentation.Slides[slideIndex].Shapes, dt);

            }
            
        }
        private static void replaceVariable(Microsoft.Office.Interop.PowerPoint.Shapes shapes, string varName, string value)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                Type t = shape.GetType();
                PropertyInfo p = t.GetProperty("GroupItems");
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.TextRange.Text.Contains("<<" + varName + ">>"))
                    {
                        shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Replace("<<" + varName + ">>", value);
                    }
                }
            }
        }
        private static void replaceVariable(Microsoft.Office.Interop.PowerPoint.Shapes shapes, System.Data.DataTable dt)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    string text = shape.TextFrame.TextRange.Text;
                    string index = Regex.Match(text, @"\d+").Value;
                    if (text.Contains("<<") && text.Contains(">>"))
                    {
                        text = text.Replace("<<", "");
                        text = text.Replace(">>", "");
                        if (dt.Columns.Contains(text))
                        {
                            shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Replace("<<" + text + ">>", dt.Rows[0][text].ToString());
                        } else if (index != "")
                        {
                            string field = text.Replace(index, "");
                            shape.TextFrame.TextRange.Text = shape.TextFrame.TextRange.Text.Replace("<<" + text + ">>", dt.Rows[Int32.Parse(index) - 1][field].ToString());
                        }
                    }
                }

            }
        }
        private static void replaceChartVariable(Microsoft.Office.Interop.PowerPoint.Shapes shapes, System.Data.DataTable dt)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in shapes)
            {
                if (shape.HasChart == MsoTriState.msoTrue)
                {
                    var chartData = shape.Chart.ChartData;
                    chartData.Activate();

                    var workbook = chartData.Workbook;
                    workbook.Application.Visible = false;
                    var dataSheet = workbook.Worksheets[1];

                    var firstColNumber = 1;
                    var firstRowNumber = 1;
                    var dataRowNumber = 2;
                    int iCol = firstColNumber;
                    foreach (DataColumn dc in dt.Columns)
                    {
                        dataSheet.Cells[firstRowNumber, iCol].Value = dc.ColumnName;
                        iCol++;
                    }
                    int iRow = dataRowNumber;
                    foreach (DataRow dr in dt.Rows)
                    {
                        iCol = 1;
                        foreach (DataColumn dc in dt.Columns)
                        {
                            dataSheet.Cells[iRow, iCol].Value = dr[dc.ColumnName];
                            iCol++;
                        }
                        iRow++;
                    }
                    workbook.Close(true);
                    shape.Chart.Refresh();
                }
            }
        }
    }
}
