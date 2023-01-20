using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Report1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Config.GetConfigurationValue();
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-EN");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            _ = StartGenReport();
            
        }
        private async Task StartGenReport()
        {
            Status.Text = "Running";
            int retry = 0;
            await Task.Run(async () =>
            {
                for (retry = 0; retry <= Config.retries; retry++)
                {
                    try
                    {
                        var dts = ExcelData.readData(textBox1.Text);
                        Powerpoint.initialReport(@textBox2.Text);
                        int numSlide = 1;
                        int i = 1;
                        int chartIdx;
                        foreach (DataTable dt in dts)
                        {
                            string processName = "";
                            string year = DateTime.Now.ToString("yyyy");
                            if (i % Config.numberOfCharts == 1)
                            {
                                processName = dt.TableName;
                                string period = "Month : " + monthCalendar1.SelectionRange.Start.ToString("MMMM yyyy");
                                Powerpoint.genReport(processName, period, year, dt, numSlide);
                            }
                            else
                            {
                                if (i % Config.numberOfCharts == 0)
                                {
                                    chartIdx = Config.numberOfCharts - 1;
                                    
                                } else
                                {
                                    chartIdx = (i % Config.numberOfCharts) - 1;
                                }
                                Powerpoint.genReport(processName, "", year, dt, numSlide, true, chartIdx);
                                if (i % Config.numberOfCharts == 0)
                                {
                                    numSlide++;
                                }
                                
                            }
                            System.Threading.Thread.Sleep(1000);
                            i++;
                        }
                        Powerpoint.CloseReport(@textBox2.Text, numSlide);
                        break;
                    }
                    catch (Exception ex)
                    {
                        if (retry > Config.retries)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                
            });
            Status.Text = "Done";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Images (*.XLS;*.XLSX;*.XLSM;*.XLM)|*.XLS;*.XLSX;*.XLSM;*.XLM|" +
                        "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (String file in openFileDialog.FileNames)
                    {
                        textBox1.Text = file;
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PowerPoint File (*.pptx;*.ppt)|*.pptx;*.ppt|" +
                        "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (String file in openFileDialog.FileNames)
                    {
                        textBox2.Text = file;
                    }
                }
            }
        }
    }
}
