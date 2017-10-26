using System;
using NetOffice.ExcelApi;

namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            CriarExcel();
            //LerExcel();
        }

        public static void CriarExcel(){
            Application ex = new Application();
            Workbook workbook = ex.Workbooks.Add();

            Worksheet workSheet = workbook.Worksheets[1] as Worksheet;

            workSheet.Name = "nome";
/*
            ex.Cells[1,1].Value = "Ford";
            ex.Cells[1,2].Value = "Focus GHIA";
            ex.Cells[1,3].Value = "2.0";
*/
            ex.ActiveWorkbook.SaveAs(@"C:\Users\41346828865\excel\lista.xls");
            ex.Quit();
        }

         public static void LerExcel(){
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\41346828865\excel\lista.xls");

            string modelo = ex.Cells[1,1].Value.ToString();
            Console.WriteLine(modelo);

            ex.Quit();
        }
    }
}
