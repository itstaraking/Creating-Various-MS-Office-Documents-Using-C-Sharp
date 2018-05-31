using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Drawing.Chart;

namespace excelFileCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                // Create worksheets
                excel.Workbook.Worksheets.Add("Worksheet1");
                excel.Workbook.Worksheets.Add("Worksheet2");
                excel.Workbook.Worksheets.Add("Worksheet3");

                List<string[]> headerRow = new List<string[]>()
                {
                    new string[] { "Maths", "Chemistry", "Physics", "English"}
                };

                string headerRange = "A1: " + Char.ConvertFromUtf32(headerRow[0].Length + 64) + 1;
                var worksheet = excel.Workbook.Worksheets["Worksheet1"];
                worksheet.Cells[headerRange].LoadFromArrays(headerRow);

                worksheet.Cells[headerRange].Style.Font.Bold = true;
                worksheet.Cells[headerRange].Style.Font.Size = 14;
                worksheet.Cells[headerRange].Style.Font.Color.SetColor(Color.Blue);

                worksheet.Cells["F1"].Value = "Sample Grade Chart" + Environment.NewLine;

                var cellData = new List<object[]>()
                {
                    new object[]{25, 30, 10, 25},
                    new object[]{90, 50, 45, 45},
                    new object[]{20, 45, 45, 50},
                    new object[]{44, 35, 66, 71},
                    new object[]{65, 43, 65, 43},

                };
                worksheet.Cells[2, 1].LoadFromArrays(cellData);

                var shape = worksheet.Drawings.AddShape("MyShape", eShapeStyle.Rect);
                shape.SetPosition(8, 0, 1, 0);
                shape.SetSize(400, 200);
                shape.Text = "This is a rectangle";

                var img = Image.FromFile("1225Logo.jpg");
                var pic = worksheet.Drawings.AddPicture("MyPicture", img);
                pic.SetPosition(10, 0, 6, 0);

                var pieChart = (ExcelPieChart)worksheet.Drawings.AddChart("crtGradePieChart", OfficeOpenXml.Drawing.Chart.eChartType.PieExploded3D);
                pieChart.SetPosition(12, 0, 6, 0);
                pieChart.SetSize(300, 300);
                pieChart.Series.Add("A2:D2", "A2:D2");
                pieChart.Title.Text = "Grade Pie Chart";

                pieChart.DataLabel.ShowCategory = true;
                pieChart.DataLabel.ShowLeaderLines = true;

                pieChart.DataLabel.ShowPercent = true;
                pieChart.Legend.Remove();

                FileInfo excelFile = new FileInfo("sample.xlsx");
                excel.SaveAs(excelFile);

                bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                if (isExcelInstalled)
                {
                    System.Diagnostics.Process.Start(excelFile.ToString());
                }
                Console.WriteLine("Program executed successfully.");
            }
        }
    }
}
