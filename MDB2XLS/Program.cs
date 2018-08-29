using MDB2XLS;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDB2XLS {
    class Program {
        static void Main(string[] args) {

            DataSet teste = Suporte.LoadFromFile(@"pathMDBSrc.mdb");
            //ExcelUtility.CreateExcel(teste, @"C:\Users\Urbgames\Desktop\exported.xlsx");
            Suporte.DataTableToExcelFile(teste.Tables[0], @"pathXLSXDest.xlsx", null);


        }
    }
}
