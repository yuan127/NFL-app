using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

// comment
namespace NFL_app
{
    class Program
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet MySheet = null;
        private static Excel.Range xlRange = null;
        static void Main(string[] args)
        {
            string startTime = DateTime.Now.ToString();
            Console.WriteLine("Start time: {0}", startTime);

            MyApp = new Excel.Application
            {
                Visible = false
            };
            string XLS_PATH = "C:\\Users\\liping.yuan\\Documents\\visual studio 2017\\Projects\\NFL app\\NFL app\\NFL_Small_Set";
            MyBook = MyApp.Workbooks.Open(XLS_PATH);
            MySheet = (Excel.Worksheet)MyBook.Sheets["1999-2013"];
            xlRange = MySheet.UsedRange;
            int lastRow = MySheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

            List<season> seasonListQ = new List<season>();
            List<season> seasonListR = new List<season>();

            string name = null;
            string position = null;
            double passingYard = 0;
            double rushingYard = 0;
            int indexQ = -1;
            int indexR = -1;

            for (int i = 2; i <= lastRow; i++)
            {               
                position = xlRange.Cells.Value2[i, 22].ToString();

                if (position == "QB")
                {
                    name = xlRange.Cells.Value2[i, 2].ToString();
                    passingYard = xlRange.Cells.Value2[i, 11];
                    season rowi = new season();

                    if (seasonListQ.Count == 0)
                    {                        
                        rowi.player = name;
                        rowi.position = position;
                        rowi.passingYard = passingYard;                        //r
                        seasonListQ.Add(rowi);
                        //Console.WriteLine("FIRST player {0}  position  {1}", name, position);
                    }
                    else
                    {
                        indexQ = seasonListQ.FindIndex(x => x.player == name);
                        if (indexQ == -1)
                        {                          
                            rowi.player = name;
                            rowi.position = position;
                            rowi.passingYard = passingYard;                           
                            seasonListQ.Add(rowi);                            
                        }
                        else
                        {                      
                            seasonListQ[indexQ].passingYard += passingYard;
                            //Console.WriteLine("UPDATE player {0} Total passing Yards: {1} ", name, seasonList[index].passingYard);
                        }
                    }
                }

                if (position == "RB")
                {
                    name = xlRange.Cells.Value2[i, 2].ToString();                  
                    rushingYard = xlRange.Cells.Value2[i, 15];
                    season rowi = new season();

                    if (seasonListR.Count == 0)
                    {                      
                        rowi.player = name;
                        rowi.position = position;                        
                        rowi.rushingYard = rushingYard;
                        seasonListR.Add(rowi);                      
                    }
                    else
                    {
                        indexR = seasonListR.FindIndex(x => x.player == name);
                        if (indexR == -1)
                        {                           
                            rowi.player = name;
                            rowi.position = position;                           
                            rowi.rushingYard = rushingYard;
                            seasonListR.Add(rowi);
                            //Console.WriteLine("player {0}  position  {1}", name, position);
                        }
                        else
                        {                           
                            seasonListR[indexR].rushingYard += rushingYard;
                            //Console.WriteLine("UPDATE player {0} Total rushing Yards: {1} ", name, seasonList[index].rushingYard);
                        }
                    }

                }

            }

            double maxPassingYards = seasonListQ.Max(x => x.passingYard);
            int maxIndexP = seasonListQ.FindIndex(x => x.passingYard == maxPassingYards);
            string maxPassingPlayer = seasonListQ[maxIndexP].player;            
            Console.WriteLine("Best Player - passing: {0} Total Yards: {1}", maxPassingPlayer, maxPassingYards);

            double maxRushingYards = seasonListR.Max(x => x.rushingYard);
            int maxIndexR = seasonListR.FindIndex(x => x.rushingYard == maxRushingYards);
            string maxRushingPlayer = seasonListR[maxIndexR].player;
            Console.WriteLine("Best Player - rushing: {0} Total Yards: {1}", maxRushingPlayer, maxRushingYards);

            //season row1 = new season();
            //row1.player = xlRange.Cells.Value2[2, 2].ToString();
            //row1.position = xlRange.Cells.Value2[2, 22].ToString();
            //row1.rushingYard = xlRange.Cells.Value2[2, 15];
            //row1.passingYard = xlRange.Cells.Value2[2, 11];

            //season row2 = new season();
            //row2.player = xlRange.Cells.Value2[3, 2].ToString();
            //row2.position = xlRange.Cells.Value2[3, 22].ToString();
            //row2.rushingYard = xlRange.Cells.Value2[3, 15];
            //row2.passingYard = xlRange.Cells.Value2[3, 11];

            string endTime = DateTime.Now.ToString();
            Console.WriteLine("End time: {0}", endTime);

            Console.ReadKey();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(MySheet);
            MyBook.Close();
            Marshal.ReleaseComObject(MyBook);
            MyApp.Quit();
            Marshal.ReleaseComObject(MyApp);

        }
    }
}
