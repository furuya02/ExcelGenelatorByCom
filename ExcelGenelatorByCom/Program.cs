using System;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Reflection;
using System.IO;

namespace ExcelGenelatorByCom {
    class Program {
        static void Main(string[] args) {


            args = new String[] { "input.csv", "output.xlsx" };

            // アプリケーションのフルパス
            var appPath = Assembly.GetExecutingAssembly().Location;
            if (args.Length != 2) {
                Console.WriteLine($"use: mono {Path.GetFileName(appPath)} input.csv output.xlsx");
                return;
            }
            var appDirectory = Path.GetDirectoryName(appPath);

            // テンプレートExcel
            var templateExcelName = Path.Combine(appDirectory, "template.xlsx");
            if (!File.Exists(templateExcelName)) {
                Console.WriteLine($"ERROR {templateExcelName} not Found.");
                return;
            }

            // 入力CSV
            var inputCsvName = Path.Combine(appDirectory, args[0]);
            if (!File.Exists(inputCsvName)) {
                Console.WriteLine($"ERROR {inputCsvName} not Found.");
                return;
            }

            // 出力Excel
            var outputExcelName = Path.Combine(appDirectory, args[1]);
            if (File.Exists(outputExcelName)) {
                File.Delete(outputExcelName);
            }

            // 入力データ
            object[,] datas = new object[29, 5]; // データを差し込む範囲は、B35～F63
            var lines = File.ReadAllLines(inputCsvName);
            foreach (var item in lines.Select((line, row) => new { line, row })) {
                var values = item.line.Split(',');
                foreach (int i in Enumerable.Range(0, 5)) {
                    datas[item.row, i] = values[i];
                }
            }

            Application excel = null;
            Workbooks workBooks = null;
            Workbook workBook = null;
            Worksheet sheet = null;
            Range range = null;
            try {
                excel = new Application();
                excel.Visible = false;
                workBooks = excel.Workbooks;
                workBook = workBooks.Open(templateExcelName);
                sheet = workBook.Sheets[1];
                range = sheet.Range["B35","F63"];
                range.Value2 = datas; // データをコピーする
                workBook.SaveAs(outputExcelName); // 出力ファイル名で保存する


                // PDF化
                var pdfName = Path.Combine(appDirectory, "sample.pdf");
                workBook.ExportAsFixedFormat(
                    XlFixedFormatType.xlTypePDF,
                    pdfName,
                    XlFixedFormatQuality.xlQualityStandard,
                    true,
                    true,
                    Type.Missing,
                    Type.Missing,
                    false,
                    Type.Missing);


                excel.Quit();
            } finally {
                Marshal.ReleaseComObject(range);
                Marshal.ReleaseComObject(sheet);
                Marshal.ReleaseComObject(workBook);
                Marshal.ReleaseComObject(workBooks);
                Marshal.ReleaseComObject(excel);
            }
        }
    }
}
