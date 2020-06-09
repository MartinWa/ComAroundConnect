using Microsoft.AnalysisServices.AdomdClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComAroundConnect
{
    class Program
    {
        static void Main(string[] args)
        {

            // Code from https://csharp.hotexamples.com/examples/-/AdomdConnection/Open/php-adomdconnection-open-method-examples.html
            using (AdomdConnection mdConn = new AdomdConnection())
            {
                mdConn.ConnectionString = "provider=msolap;Data Source=(local);initial catalog=HabraCube;";
                mdConn.Open();
                AdomdCommand mdCommand = mdConn.CreateCommand();
                mdCommand.CommandText = "SELECT {[Measures].[Vote], [Measures].[Votes Count]} ON COLUMNS, [Dim Time].[Month Name].MEMBERS ON ROWS FROM [Habra DW]"; // << MDX Query // work with CellSet 
                CellSet cs = mdCommand.ExecuteCellSet(); 
                if (cs.Axes.Count != 2) return;
                TupleCollection tuplesOnColumns = cs.Axes[0].Set.Tuples;
                TupleCollection tuplesOnRows = cs.Axes[1].Set.Tuples; 
                Console.Write("{0,-12}", "Item");
                for (int col = 0; col < tuplesOnColumns.Count; col++)
                {
                    Console.Write("{0,-12}", tuplesOnColumns[col].Members[0].Caption);
                }
                Console.WriteLine();
                for (int row = 0; row < tuplesOnRows.Count; row++)
                {
                    Console.Write("{0,-12}", tuplesOnRows[row].Members[0].Caption);
                    for (int col = 0; col < tuplesOnColumns.Count; col++)
                    { Console.Write("{0,-12}", cs.Cells[col, row].Value); }
                    Console.WriteLine();
                }
                Console.ReadLine();
            }
        }
    }
}
