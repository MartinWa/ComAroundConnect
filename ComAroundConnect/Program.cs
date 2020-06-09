using Microsoft.AnalysisServices.AdomdClient;
using System;

namespace ComAroundConnect
{
    class Program
    {
        static void Main()
        {
            const string Server = "";
            const string UserName = "";
            const string Password = "";
            const string Model = "";


            // Code from https://csharp.hotexamples.com/examples/-/AdomdConnection/Open/php-adomdconnection-open-method-examples.html
            using (AdomdConnection mdConn = new AdomdConnection())
            {
                mdConn.ConnectionString = $"Provider=MSOLAP;Data Source=asazure://westeurope.asazure.windows.net/{Server};Initial Catalog={Model};User ID={UserName};Password={Password};Persist Security Info=True;Impersonation Level=Impersonate";
                mdConn.Open();
                AdomdCommand mdCommand = mdConn.CreateCommand();

                mdCommand.CommandText = @"
                    SELECT
                        NON EMPTY[DimDate].[Year].members ON 0,
                        [Measures].[Votes] ON 1
                    FROM[Model]"; // << MDX Query 
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
