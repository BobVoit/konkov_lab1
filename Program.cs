using OfficeOpenXml;
using System;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace MyApp // Note: actual namespace depends on the project name.
{
    internal class Program
    {
        // private static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Workers.mdb;";
        private static string connectString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;";
        private OleDbConnection myConnection;

        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            using (var package = new ExcelPackage(@"D:\busines\konkov_lab1\konkov_lab1\excel_data.xlsx"))
            {
                var sheets = package.Workbook.Worksheets;

                OleDbConnectionStringBuilder bldr = new OleDbConnectionStringBuilder();
                bldr.DataSource = "../../../Database.accdb";
                bldr.PersistSecurityInfo = true;
                bldr.Provider = "Microsoft.ACE.OLEDB.12.0";
                // string connectString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Database.mdb;";
                // OleDbConnection dbConnection = new OleDbConnection(connectString);

                OleDbConnection dbConnection = new OleDbConnection(bldr.ConnectionString);

                foreach (ExcelWorksheet currentPage in sheets)
                {
                    string pageName = currentPage.Name;

                    foreach (var cur in currentPage.Columns)
                    {
                        var rr = cur;
                    }

                    ExcelRange cells = currentPage.Cells;

                    int countRows = currentPage.Dimension.End.Row;
                    int countColumns = currentPage.Dimension.End.Column;

                    List<string> columns = new List<string>();

                    for (int col = 1; col <= countColumns; col++)
                    {
                        columns.Add(Convert.ToString(cells[1, col].Value));
                    }

                    string queryText = "CREATE TABLE " + pageName + "(";

                    bool isFirst = true;
                    string columnsText = "";
                    foreach (string column in columns)
                    {
                        if (!isFirst)
                        {
                            queryText += ", ";
                            columnsText += ", ";
                        }
                        else
                            isFirst = false;
                        queryText += column + " CHAR";
                        columnsText += column;
                    }

                    queryText += ");";

                    using (dbConnection = new OleDbConnection(bldr.ConnectionString))
                    {
                        try
                        {
                            dbConnection.Open();
                            OleDbCommand cnd = new OleDbCommand(queryText, dbConnection);
                            OleDbDataReader rdr = cnd.ExecuteReader();
                        }
                        catch (SqlException ex)
                        {
                            Console.WriteLine(ex.Message);
                            continue;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            continue;
                        }
                    }
                 
                    for (int row = 2; row <= countRows; row++) 
                    {
                        queryText = "INSERT INTO " + pageName + " VALUES ";
                        queryText += "(";
                        bool isFirstVal = true;
                        for (int col = 1; col <= countColumns; col++)
                        {
                            if (!isFirstVal)
                            {
                                queryText += ", ";
                            }
                            else
                                isFirstVal = false;
                            queryText += "'" +  Convert.ToString(cells[row, col].Value) + "'";
                        }
                        queryText += ");";

                        using (dbConnection = new OleDbConnection(bldr.ConnectionString))
                        {
                            try
                            {
                                dbConnection.Open();
                                OleDbCommand cnd = new OleDbCommand(queryText, dbConnection);
                                OleDbDataReader rdr = cnd.ExecuteReader();
                            }
                            catch (SqlException ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }
            }
        }
    }
}