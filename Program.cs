using System;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace Reporting
{   
    public class Program {

        static void Main(String[] args){
            var report = new Program();
            Console.WriteLine(report.ActualReport());
            Console.WriteLine(report.ExpectedReport());
        }

        public object ActualReport(){
            // load from excel
            var actualFilePath = @"/Users/moorel/Desktop/Projects/C#/O2/Reporting/actual/report_actual.xlsm";
            FileInfo actualFile = new FileInfo(actualFilePath);
 
            using (ExcelPackage actualPackage = new ExcelPackage(actualFile)){   
            ExcelWorksheet worksheet = actualPackage.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;
                int ColCount = worksheet.Dimension.Columns;
                    var rawText = string.Empty;
                    for (int row = 1; row <= rowCount; row++){
                     for (int col = 1; col <= ColCount; col++){   
                        rawText += worksheet.Cells[row, col].Value.ToString() + "\t";    
                        }
                    rawText+="\r\n";
                    }
                return rawText;
            }
        }

        public object ExpectedReport() {
            // load from excel
            var expectedFilePath = @"/Users/moorel/Desktop/Projects/C#/O2/Reporting/expected/report_expected.xlsm";
            FileInfo expectedFile = new FileInfo(expectedFilePath);
 
            using (ExcelPackage expectedPackage = new ExcelPackage(expectedFile)){   
            ExcelWorksheet worksheet = expectedPackage.Workbook.Worksheets[1];
                int rowCount = worksheet.Dimension.Rows;
                int ColCount = worksheet.Dimension.Columns;
                    var rawText = string.Empty;
                    for (int row = 1; row <= rowCount; row++){
                    for (int col = 1; col <= ColCount; col++){   
                        rawText += worksheet.Cells[row, col].Value.ToString() + "\t";    
                        }
                    rawText+="\r\n";
                    }
                return rawText;;
            }
        }
    }
}
