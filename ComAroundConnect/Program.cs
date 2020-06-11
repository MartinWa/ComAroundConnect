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

            if (string.IsNullOrEmpty(Server) || string.IsNullOrEmpty(UserName) || string.IsNullOrEmpty(Password) || string.IsNullOrEmpty(Model))
            {
                Console.WriteLine("You must set server/username/password/model");
            }
            else
            {
                var connectionString = $"Provider=MSOLAP;Data Source=asazure://westeurope.asazure.windows.net/{Server};Initial Catalog={Model};User ID={UserName};Password={Password};Persist Security Info=True;Impersonation Level=Impersonate";
                ListVotesPerYear(connectionString);
                RetrieveCubesAndDimensions(connectionString);
            }
            Console.WriteLine("Press any key");
            Console.ReadLine();
        }

        private static void ListVotesPerYear(string connectionString)
        {
            // Code from https://csharp.hotexamples.com/examples/-/AdomdConnection/Open/php-adomdconnection-open-method-examples.html
            using (AdomdConnection mdConn = new AdomdConnection())
            {
                mdConn.ConnectionString = connectionString;
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
                mdConn.Close();
            }
        }

        private static void RetrieveCubesAndDimensions(string connectionString)
        {
            // From https://docs.microsoft.com/en-us/analysis-services/adomd/multidimensional-models-adomd-net-client/retrieving-metadata-working-with-adomd-net-object-model?view=asallproducts-allversions

            //Connect to the local server
            using (AdomdConnection conn = new AdomdConnection(connectionString))
            {
                conn.Open();

                //Loop through every cube
                foreach (CubeDef cube in conn.Cubes)
                {
                    //Skip hidden cubes.
                    if (cube.Name.StartsWith("$"))
                        continue;

                    //Write the cube name
                    Console.WriteLine(cube.Name);

                    //Write out all dimensions, indented by a tab.
                    foreach (Dimension dim in cube.Dimensions)
                    {
                        Console.WriteLine("\t");
                        Console.WriteLine(dim.Name);
                    }
                }

                //Close the connection
                conn.Close();
            }
        }
    }
}
