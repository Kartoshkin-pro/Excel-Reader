
using Microsoft.Office.Interop.Excel;

class Program
{
    static void Main(string[] args)
    {
        //Create COM Objects.
        Application excelApp = new Application();

        if (excelApp == null)
        {
            Console.WriteLine("Excel is not installed!!");
            return;
        }
        string[] path = new string[2];

        path[0] = @"C:\Users\1\Desktop\NmarketTestTask\NmarketTestTask\Files\Excel\1.xlsx";
        path[1] = @"C:\Users\1\Desktop\NmarketTestTask\NmarketTestTask\Files\Excel\2.xlsx";
        

        for(int num = 0; num <= 2; num++)
        {
            Workbook excelBook = excelApp.Workbooks.Open(path[num]);
            _Worksheet excelSheet = excelBook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

            int rows = excelRange.Rows.Count;
            int cols = excelRange.Columns.Count;

            for (int i = 1; i <= rows; i++)
            {
                //create new line
                Console.Write("\r\n");
                for (int j = 1; j <= cols; j++)
                {

                    //write the console
                    if (excelRange.Cells[i, j] != null && excelRange.Cells[i, j].Value2 != null)
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
                }
            }
            //after reading, relaase the excel project
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

        }

        Console.ReadLine();
    }
}
