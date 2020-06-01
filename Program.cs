using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using Microsoft.Office.Interop;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace DataExportApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get all data from the Source database 
            DataSet ds = GetAllDataFromSource();

            //Export all data to Excel
            ExportDataSetToExcel(ds);

            Console.WriteLine("Done");
            Console.ReadLine();
        }

        private static DataSet GetAllDataFromSource()
        {
            //Declare
            System.Data.DataTable userTables = null;
            OleDbDataAdapter oledbAdapter;
            DataSet ds = new DataSet();
            List<string> tableNames = new List<string>();
            using (OleDbConnection myConnection = new OleDbConnection())
            {
                myConnection.ConnectionString = ConfigurationManager.ConnectionStrings["SourceDatabaseConnectionString"].ConnectionString;
                //Connect to Source database
                myConnection.Open();

                //Restrict the GetSchema() to return "Tables" schema information only.
                string[] restrictions = new string[4];
                restrictions[3] = "Table";
                userTables = myConnection.GetSchema("Tables", restrictions);

                for (int i = 0; i < userTables.Rows.Count; i++)
                {
                    var tableName = userTables.Rows[i][2].ToString();
                    oledbAdapter = new OleDbDataAdapter($"select * from {tableName}", myConnection);
                    oledbAdapter.Fill(ds, $"{tableName}");

                    if (ds.Tables[$"{tableName}"].Rows.Count > 0)
                    {
                        Console.WriteLine("Rows: " + ds.Tables[$"{tableName}"].Rows.Count);
                    }
                    oledbAdapter.Dispose();

                }
                myConnection.Close();
            }
            return ds;
        }

        /// <summary>
        /// This method takes DataSet as input parameter and it exports the same to excel
        /// </summary>
        /// <param name="ds"></param>
            private static void ExportDataSetToExcel(DataSet ds)
            {
                //Creae an Excel application instance
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                //Create an Excel workbook instance and open it from the predefined location

                foreach (System.Data.DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
                    excelWorkSheet.Name = table.TableName;

                    //Columns
                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    }

                    //Rows
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            try
                            {
                                excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                            }
                            catch(Exception ex)
                            {
                                Console.WriteLine($"Error in table: {excelWorkSheet.Name} - Cells - j: {j}, k:{k}, data: {table.Rows[j].ItemArray[k].ToString()}");
                                Console.WriteLine(ex);                            
                            }                        
                        }
                    }
                }
                string fileName = System.IO.Path.Combine(System.Configuration.ConfigurationManager.AppSettings["TargetDirectory"], $@"test-{DateTime.Now.ToString("yyyyMMddHHmmss")}.xls");

                excelWorkBook.SaveAs(fileName);
                excelWorkBook.Close();
                excelApp.Quit();

            }
    }
}
