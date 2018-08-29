using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MDB2XLS {
    public class Suporte {


        ///

        /// Converts excel file sheet content into DataTable.
        ///

        /// 
        /// 
        /// 
        public static DataTable ExcelFileToDataTable(string p_fileUrl, int p_sheetIndex) {
            DataTable dbSchema = new DataTable();
            DataSet dataSet = new DataSet();
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + p_fileUrl + ";" + "Extended Properties=Excel 8.0;Mode=Read;";

            using (OleDbConnection conn = new OleDbConnection(strConn)) {
                conn.Open();

                // Get all sheetnames from an excel file into data table
                dbSchema = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                string sheetName = dbSchema.Rows[p_sheetIndex]["TABLE_NAME"].ToString();
                string selectCommandText = string.Format("SELECT * FROM [{0}]", sheetName);

                using (OleDbDataAdapter adapter = new OleDbDataAdapter(selectCommandText, conn)) {
                    adapter.Fill(dataSet);
                }

                return dataSet.Tables[0];
            }
        }

        ///

        /// Converts DataTable to Excel-File
        ///

        /// <summary>
        /// Loads a Microsoft Access Database file into a DataSet object.
        /// The file can be the in the newer ACCDB format or MDB legacy format.
        /// </summary>
        /// <param name="fileName">The file name to load.</param>
        /// <returns>A DataSet object with the Tables object populated with the contents of the specified Microsoft Access Database.</returns>
        public static DataSet LoadFromFile(string fileName) {
            DataSet result = new DataSet();

            // For convenience, the DataSet is identified by the name of the loaded file (without extension).
            result.DataSetName = Path.GetFileNameWithoutExtension(fileName).Replace(" ", "_");

            // Compute the ConnectionString (using the OLEDB v12.0 driver compatible with ACCDB and MDB files)
            fileName = Path.GetFullPath(fileName);
            string connString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}", fileName);

            // Opening the Access connection
            using (OleDbConnection conn = new OleDbConnection(connString)) {
                conn.Open();

                // Getting all user tables present in the Access file (Msys* tables are system thus useless for us)
                DataTable dt = conn.GetSchema("Tables");
                List<string> tablesName = dt.AsEnumerable().Select(dr => dr.Field<string>("TABLE_NAME")).Where(dr => !dr.StartsWith("MSys")).ToList();

                // Getting the data for every user tables
                foreach (string tableName in tablesName) {
                    using (OleDbCommand cmd = new OleDbCommand(string.Format("SELECT * FROM [{0}]", tableName), conn)) {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd)) {
                            // Saving all tables in our result DataSet.
                            DataTable buf = new DataTable("[" + tableName + "]");
                            adapter.Fill(buf);
                            result.Tables.Add(buf);
                        } // adapter
                    } // cmd
                } // tableName
            } // conn

            // Return the filled DataSet
            return result;
        }

        /// 
        /// 
        /// 
        /// 
        public static bool DataTableToExcelFile(DataTable p_dt, string p_filePath, EventHandler p_onRowHandled) {
            try {
                // fix table name
                if (p_dt.TableName == "" || p_dt.TableName.Equals("Table", StringComparison.OrdinalIgnoreCase) == true) {
                    p_dt.TableName = "NONAME";
                }

                // fix file path name
                if (p_filePath.EndsWith(".xlsx") == false) {
                    p_filePath += ".xlsx";
                }

                // insure file not exist!
                System.IO.FileInfo fi = new System.IO.FileInfo(p_filePath);
                if (fi.Name.Replace(".xlsx", "").Trim().Length <= 0) {
                    p_filePath = p_filePath.Replace(".xlsx", p_dt.TableName + ".xlsx");
                }
                if (System.IO.File.Exists(p_filePath)) {
                    System.IO.File.Delete(p_filePath);
                }

                using (OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + p_filePath + ";Extended Properties=Excel 12.0 Xml;")) {
                    using (OleDbCommand command = new OleDbCommand()) {
                        command.Connection = con;
                        con.Open();

                        // create table 
                        command.CommandText = GenerateSqlStatementCreateTable(p_dt);
                        command.ExecuteNonQuery();

                        DataRow dr;
                        OleDbParameter oleDbParameter;

                        // Create columns & values (as parameters) script
                        string columns, parameters;
                        GenerateColumnsString(p_dt, out columns, out parameters);

                        for (int i = 0; i < p_dt.Rows.Count; i++) {
                            dr = p_dt.Rows[i];
                            Console.WriteLine(i + "/" + p_dt.Rows.Count);
                            // set insert statement parameters
                            command.Parameters.Clear();
                            for (int j = 0; j < p_dt.Columns.Count; j++) {
                                oleDbParameter = new OleDbParameter();
                                oleDbParameter.ParameterName = "@p" + j;

                                if (dr.IsNull(j) == true) {
                                    oleDbParameter.Value = DBNull.Value;
                                }
                                else {
                                    oleDbParameter.Value = dr[j];
                                }
                                command.Parameters.Add(oleDbParameter);
                            }

                            command.CommandText = string.Format("INSERT INTO {0} ({1}) VALUES ({2})", p_dt.TableName, columns, parameters);
                            command.ExecuteNonQuery();

                            // Perform step...
                            if (p_onRowHandled != null) {
                                p_onRowHandled(p_dt, EventArgs.Empty);
                            }
                        }
                    }

                    con.Close();
                }
                return true;
            }
            catch (Exception ex) {
                throw (ex);
            }
        }


        ///

        /// Generates columns names string delimited by commas
        /// also supply parameters for the given columns, for example:
        /// 
        /// out string p_columns = [columnname0],[columnname1],[columnname2]
        /// out string p_params = @p0, @p1, @p2
        /// 
        ///

        /// 
        /// 
        /// 
        private static void GenerateColumnsString(DataTable p_dt, out string p_columns, out string p_params) {
            StringBuilder sbColumns = new StringBuilder();
            StringBuilder sbParams = new StringBuilder();
            for (int i = 0; i < p_dt.Columns.Count; i++) {
                if (i != 0) {
                    sbColumns.Append(',');
                    sbParams.Append(',');
                }
                sbColumns.AppendFormat("[{0}]", p_dt.Columns[i].ColumnName);
                sbParams.AppendFormat("@p{0}", i);
            }

            p_columns = sbColumns.ToString();
            p_params = sbParams.ToString();
        }

        ///

        /// Create SQL-Script for creating table which represent the given DataTable
        ///

        /// 
        /// 
        private static string GenerateSqlStatementCreateTable(DataTable p_dt) {
            StringBuilder sbCreateTable = new StringBuilder();

            DataColumn dc;
            sbCreateTable.AppendFormat("CREATE TABLE {0} (", p_dt.TableName);
            for (int i = 0; i < p_dt.Columns.Count; i++) {
                dc = p_dt.Columns[i];
                if (i != 0) {
                    sbCreateTable.Append(",");
                }

                string dataType = dc.DataType.Equals(typeof(double)) ? "DOUBLE" : "NVARCHAR";

                sbCreateTable.AppendFormat("[{0}] {1}", dc.ColumnName, dataType);
            }
            sbCreateTable.Append(")");

            return sbCreateTable.ToString();
        }

    }
}
