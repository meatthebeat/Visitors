using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace ConsoleApplication1
{
    class Program
    {
        public class Visitors
        {
            public int number { get; set; }
            public string Name { get; set; }
            public string email { get; set; }
            public int ID { get; set; }
            public bool day1 { get; set; }
            public bool day2 { get; set; }
            public bool day3 { get; set; }
            public int cost { get; set; }
            public bool tent { get; set; }
            public int addcost { get; set; }
            public int total { get; set; }


        }

        public static string RandomString(int size)
        {
            StringBuilder builder = new StringBuilder();
            Random random = new Random();
            char ch;
            for (int i = 0; i < size; i++)
            {

                ch = Convert.ToChar(Convert.ToInt32(Math.Floor(26 * random.NextDouble() + 65)));
                builder.Append(ch);
            }
            return builder.ToString();
        }

        static void DisplayInExcel(IEnumerable<Visitors> visitors)
        {
            var excelApp = new Excel.Application();

            excelApp.Visible = true;

            excelApp.Workbooks.Add();

            Excel._Worksheet workSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            workSheet.Cells[1, "A"] = "FESTIVAL VISITORS";
            workSheet.Cells[2, "A"] = "Date of the beginning";
            workSheet.Cells[3, "A"] = "Date of the end";
            workSheet.Cells[4, "A"] = "№";
            workSheet.Cells[4, "B"] = "Name";
            workSheet.Cells[4, "C"] = "e-mail";
            workSheet.Cells[4, "D"] = "ID";
            workSheet.Cells[4, "E"] = "Day 1 (10$)";
            workSheet.Cells[4, "F"] = "Day 2 (15$)";
            workSheet.Cells[4, "G"] = "Day 3 (20$)";
            workSheet.Cells[4, "H"] = "Cost of tisket";
            workSheet.Cells[4, "I"] = "Tent(5$)";
            workSheet.Cells[4, "J"] = "Cost of add servises";
            workSheet.Cells[4, "K"] = "Total sum";


            var row = 4;
            foreach (var acct in visitors)
            {
                row++;
                workSheet.Cells[row, "A"] = acct.number;
                workSheet.Cells[row, "B"] = acct.Name;
                workSheet.Cells[row, "C"] = acct.email;
                workSheet.Cells[row, "D"] = acct.ID;
                workSheet.Cells[row, "E"] = acct.day1;
                workSheet.Cells[row, "F"] = acct.day2;
                workSheet.Cells[row, "G"] = acct.day3;
                workSheet.Cells[row, "H"] = acct.cost;
                workSheet.Cells[row, "I"] = acct.tent;
                workSheet.Cells[row, "J"] = acct.addcost;
                workSheet.Cells[row, "K"] = acct.total;

            }

            workSheet.Columns[1].AutoFit();
            workSheet.Columns[2].AutoFit();
            workSheet.Columns[3].AutoFit();
            workSheet.Columns[4].AutoFit();
            workSheet.Columns[5].AutoFit();
            workSheet.Columns[6].AutoFit();
            workSheet.Columns[7].AutoFit();
            workSheet.Columns[8].AutoFit();
            workSheet.Columns[9].AutoFit();
            workSheet.Columns[10].AutoFit();
            workSheet.Columns[11].AutoFit();

        }

        static void Main(string[] args)
        {

            var festguests = new List<Visitors>();

            int r;


            //Добавление 100 пользователей в таблицу

            //for(int i = 0 ; i < 100; i++)
            //{
            //    Random random = new Random();
            //    r = random.Next(5, 10);

            //    string s = RandomString(r);
            //    string e = s + "@gmail.com";
            //    Visitors v = new Visitors
            //    {
            //        number = i+1,
            //        Name = s,
            //        email = e,
            //        ID = 1000 + i+1,
            //        day1 = true,
            //        day2 = true,
            //        day3 = true,
            //        cost = 45,
            //        tent = true,
            //        addcost = 0,
            //        total = 45
            //    };

            //    festguests.Add(v);

            // }

            DisplayInExcel(festguests);
        }
    };
}
