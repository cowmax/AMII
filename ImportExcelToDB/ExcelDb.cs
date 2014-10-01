using System;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.Common;
using System.Web;
using System.Web.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace ImportExcelToDB
{
    /// <summary>
    /// 类：把一个Excel 表格作为一个数据表，通过 OleDb 访问
    /// </summary>
    class ExcelDb
    {
        DataSet coreDataSet;
        OleDbConnection xlsConnection;
        OleDbDataAdapter xlsDbAdapter;
        string xlsFilePath;
        string xlsSheetName;
        string oleDbTableName;
        int rowCount = 0;

        public ExcelDb(DataTable dtTable)
        {
            this.coreDataSet = new DataSet();
            this.coreDataSet.Tables.Add(dtTable);
            this.oleDbTableName = dtTable.TableName;
        }

        /// <summary>
        /// 构造函数：从Excel文件的指定Sheet 中加载数据到 DataSet 中
        /// </summary>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="sheetName">Excel 中的 Sheet 名称</param>
        /// <param name="srcTableName">数据库集中相应的 Table名称，缺省名称与Sheet 相同</param>
        public ExcelDb(string xlsFilePath, string sheetName)
        { 
            this.xlsFilePath = xlsFilePath;
            this.xlsSheetName = sheetName;

            this.OpenXlsDbConnection(xlsFilePath);
            this.GetExcelDbAdapter(sheetName);
            this.rowCount = this.FillSheetToDateSet();
        }

        #region -- BEGIN: Private functions

        private bool isSheetExist(Excel.Sheets sheets, string shtName)
        {
            bool bExist = false;
            foreach (Excel.Worksheet sht in sheets)
            {
                if (sht.Name == shtName)
                {
                    bExist = true;
                    break;
                }
            }

            return bExist;
        }

        // Get work-sheet with specified name, if the sheet do not exist, create it where bCreate is true
        private Excel.Worksheet getSheet(Excel.Sheets sheets, string shtName, bool bCreate = false)
        {
            Excel.Worksheet sht = null;
            foreach (Excel.Worksheet s in sheets)
            {
                if (s.Name == shtName)
                {
                    sht = s;
                    break;
                }
            }

            if (sht == null && bCreate)
            {
                sht = sheets.Add();
                sht.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }

            return sht ;
        }

        private string GetCreateSheetSql(string tblName, string[] columnNames)
        {
            string sqlCreateTmpl = "CREATE TABLE {0} ({1})";

            string strColumns = null;
            foreach (string colName in columnNames)
            {
                strColumns += "[" + colName + "] VarChar,";
            }

            strColumns = strColumns.Substring(0, strColumns.Length -1);

            return string.Format(sqlCreateTmpl, tblName, strColumns);
        }

        private string GetTableName(string sheetName){
            return sheetName;
        }

        // Open specified Excel file with OLEDB 
        private OleDbConnection OpenXlsDbConnection(string xlsFilePath, bool forExport = false)
        {
            if (this.xlsConnection == null)
            {
                string strConnTmpl = null;
                if (forExport)
                {
                    // 打开 Excel 文件用于创建、写入操作时，参数 IMEX = 1 会导致访问异常
                    strConnTmpl = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;'";
                }
                else
                {
                    // NOTE: Different version Excel file need specified connection string
                    // Excel 2003
                    // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + xlsFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                    // Excel 2007
                    strConnTmpl = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";
                }

                //连接数据源      
                this.xlsConnection = new OleDbConnection(string.Format(strConnTmpl, xlsFilePath));
                this.xlsConnection.Open();
            }

            return this.xlsConnection;
        }

        private void CloseXlsDbConnection()
        {
            if (this.xlsConnection != null)
            {
                this.xlsConnection.Close();
                this.xlsConnection = null;
                
                this.xlsDbAdapter.Dispose();
                this.xlsDbAdapter = null;

                this.coreDataSet.Dispose();
                this.coreDataSet = null;
            }
        }

        // Fill excel sheet data to DateSet
        private int FillSheetToDateSet(string tblName = null)
        {
            if (tblName != null)
            {
                this.oleDbTableName = tblName;
            }
            else
            {
                this.oleDbTableName = this.xlsSheetName;
            }

            this.coreDataSet = new DataSet();
            // Fill data of the table in adapter to specified DataSet with given name
            return this.xlsDbAdapter.Fill(this.coreDataSet, this.xlsSheetName);
        }

        private string GetSelectAllSql(string sheetName = null)
        {
            if (sheetName != null)
            {
                this.xlsSheetName = sheetName;
            }
            return ("select * from  [" + this.xlsSheetName + "$] ");
        }

        // Return an OleDB Adapter for excel file's specified sheet
        private OleDbDataAdapter GetExcelDbAdapter(string sheetName)
        {
            if (this.xlsDbAdapter == null)
            {
                string strSql = GetSelectAllSql(sheetName);
                this.xlsDbAdapter = new OleDbDataAdapter(strSql, this.xlsConnection);
            }

            return this.xlsDbAdapter;
        }

        private int Fill(DataSet trgDataSet, string srcTable)
        {
            return this.xlsDbAdapter.Fill(trgDataSet, oleDbTableName);
        }

        private int Export(string filepath, string strSheetName)
        {
            this.OpenXlsDbConnection(filepath, true);
            this.GetExcelDbAdapter(strSheetName);
            DataTable dt = this.GetOleDbTable();
            string tblName = this.oleDbTableName;
            int nRow = 0;

            try
            {
                using (OleDbConnection DbConn = this.xlsConnection)
                {
                    StringBuilder strSQL = new StringBuilder();
                    strSQL.Append("CREATE TABLE ").Append("[" + tblName + "]");
                    strSQL.Append("(");
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        strSQL.Append("[" + dt.Columns[i].ColumnName + "] text,");
                    }
                    strSQL = strSQL.Remove(strSQL.Length - 1, 1);
                    strSQL.Append(")");

                    OleDbCommand cmd = new OleDbCommand(strSQL.ToString(), DbConn);
                    cmd.ExecuteNonQuery();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        strSQL.Clear();
                        StringBuilder strfield = new StringBuilder();
                        StringBuilder strvalue = new StringBuilder();
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            strfield.Append("[" + dt.Columns[j].ColumnName + "]");
                            strvalue.Append("'" + dt.Rows[i][j].ToString() + "'");
                            if (j != dt.Columns.Count - 1)
                            {
                                strfield.Append(",");
                                strvalue.Append(",");
                            }
                        }
                        cmd.CommandText = strSQL.Append(" insert into [" + tblName + "]( ")
                            .Append(strfield.ToString())
                            .Append(") values (").Append(strvalue).Append(")").ToString();
                        nRow = cmd.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return nRow;
        }

        #endregion -- END :private functions

        public void Close()
        {
            CloseXlsDbConnection();
        }

        public DataTable GetOleDbTable()
        {
            return this.coreDataSet.Tables[this.oleDbTableName];
        }

        /// <summary>
        /// 函数：使用 Office.Interop.Excel 组件将DateSet 中的Table 保存到 Excel文件的 Sheet 中
        /// </summary>
        /// <param name="srcTableName">DateSet 中 Table 的名称</param>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="sheetName">Excel 中 Sheet 的名称</param>
        /// <returns></returns>
        public int SaveToXls(string xlsFilePath, string sheetName)
        {
            DataTable dtTable = this.coreDataSet.Tables[this.oleDbTableName];
            int rowCount = dtTable.Rows.Count;

            // NOTE : Must Add reference Microsoft.Office.Interop.Excel assembly
            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook xlsWrkBook = xlsApp.Application.Workbooks.Open(xlsFilePath);
            Excel.Worksheet wrkSheet = getSheet(xlsWrkBook.Sheets, sheetName, true);

            wrkSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            int rowIndex = 1;
            int colIndex = 0;

            // Fill table column name to work-sheet
            foreach (DataColumn col in dtTable.Columns)
            {
                colIndex++;
                wrkSheet.Cells[rowIndex, colIndex] = col.ColumnName;
            }

            // Fill data-row to work-sheet
            foreach (DataRow row in dtTable.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in dtTable.Columns)
                {
                    colIndex++;
                    wrkSheet.Cells[rowIndex, colIndex] = "'" + row[col.ColumnName].ToString();
                }
            }

            //设置禁止弹出保存和覆盖的询问提示框
            xlsApp.DisplayAlerts = false;
            xlsApp.AlertBeforeOverwriting = false;

            // 保存工作簿
            xlsWrkBook.Save(); // or : xlsApp.ActiveWorkbook.Save();

            // 保存excel文件 -- 似乎没有这个必要
            // xlsApp.SaveWorkspace(xlsFilePath);

            xlsApp.Quit();
            xlsApp = null;
            GC.Collect();//垃圾回收

            return rowCount;
        }

        /// <summary>
        /// 函数：使用OLEDB 将DateSet 中的Table 保存到 Excel文件的 Sheet 中
        /// </summary>
        /// <param name="srcTableName">DateSet 中 Table 的名称</param>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中 Sheet 的名称</param>
        /// <returns></returns>
        public int SaveToXlsFile(string xlsFilePath, string sheetName)
        {
            return Export(xlsFilePath, sheetName);
        }

        public int SaveToDb(DataAdapter dtAdapter, DataTable trgDbTable)
        {
            DataRow newRow = null;
            DataTable srcTable = GetOleDbTable();

            foreach (DataRow srcRow in srcTable.Rows)
            {
                newRow = trgDbTable.NewRow();

                // set each field value
                foreach (DataColumn col in trgDbTable.Columns)
                {
                    newRow[col.Caption] = srcRow[col.Caption];
                }

                trgDbTable.Rows.Add(newRow);
            }

            return dtAdapter.Update(trgDbTable.DataSet);
        }

/*
        protected void Button1_Click(object sender, EventArgs e)
        {
            string fileName = null;
            try
            {
                Boolean fileOK = false;
                String path = Server.MapPath("./doc/");
                if (FileUpload2.HasFile)
                {
                    String fileExtension =
                        System.IO.Path.GetExtension(FileUpload2.FileName).ToLower();
                    String[] allowedExtensions = { ".xls" };     //C#读取Excel中数据
                    for (int i = 0; i < allowedExtensions.Length; i++)
                    {
                        if (fileExtension == allowedExtensions[i])
                        {
                            fileOK = true;
                        }
                    }
                }

                if (fileOK)
                {
                    fileName = "r_" + DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss") + "_" + DateTime.Now.Millisecond +
                         System.IO.Path.GetExtension(FileUpload2.FileName).ToLower();
                    if (File.Exists(path + fileName))
                    {
                        Random rnd = new Random(10000);
                        fileName = fileName + rnd.Next();
                    }

                    FileUpload2.PostedFile.SaveAs(path
                        + fileName);


                }
                else
                {

                }
            }
            catch (Exception exp)
            {
            }
            ExcelToDS(Server.MapPath(".") + "\\doc\\" + fileName);
        }  
 */
    }
}
