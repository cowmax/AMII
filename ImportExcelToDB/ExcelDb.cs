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
        DataSet xlsDataSet;
        OleDbConnection xlsConnection;
        OleDbDataAdapter xlsDbAdapter;
        string xlsFilePath;
        string xlsSheetName;
        string dataTableName;

        public ExcelDb(DataTable dtTable)
        {
            this.xlsDataSet = new DataSet();
            this.xlsDataSet.Tables.Add(dtTable);
            this.dataTableName = dtTable.TableName;
            this.xlsSheetName = dtTable.TableName;
        }

        /// <summary>
        /// 构造函数：从Excel文件的指定Sheet 中加载数据到 DataSet 中
        /// </summary>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中的 Sheet 名称</param>
        /// <param name="srcTableName">数据库集中相应的 Table名称，缺省名称与Sheet 相同</param>
        public ExcelDb(string xlsFilePath, string strSheetName, string srcTableName = null)
        { 
            this.xlsFilePath = xlsFilePath;
            this.xlsSheetName = strSheetName;

            if (srcTableName == null)
            {
                this.dataTableName = GetTableName(strSheetName);
            }
            else
            {
                this.dataTableName = srcTableName;
            }
        }

        // Open the excel file and load 
        public int Open(string sheetName = null)
        {
            this.OpenXlsDbConnection(this.xlsFilePath);
            this.GetExcelDbAdapter(sheetName);
            int rowCount = this.FillSheetToDateSet();

            return rowCount;
        }

        public void Close()
        {
            CloseXlsDbConnection();
        }

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

        /// <summary>
        /// 函数：将DateSet 中的Table 保存到 Excel文件的 Sheet 中
        /// </summary>
        /// <param name="srcTableName">DateSet 中 Table 的名称</param>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中 Sheet 的名称</param>
        /// <returns></returns>
        public int SaveToXls(string xlsFilePath, string strSheetName)
        {
            DataTable dtTable = this.xlsDataSet.Tables[this.dataTableName];
            int rowCount = dtTable.Rows.Count;

            // NOTE : Must Add reference Microsoft.Office.Interop.Excel assembly
            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook xlsWrkBook= xlsApp.Application.Workbooks.Open(xlsFilePath);
            Excel.Worksheet wrkSheet = getSheet(xlsWrkBook.Sheets, strSheetName, true);

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
        /// 函数：将DateSet 中的Table 保存到 Excel文件的 Sheet 中
        /// </summary>
        /// <param name="srcTableName">DateSet 中 Table 的名称</param>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中 Sheet 的名称</param>
        /// <returns></returns>
        public int SaveToXlsFile(string xlsFilePath, string strSheetName)
        {
            DataTable dtTable = this.xlsDataSet.Tables[this.dataTableName];
            int rowCount = dtTable.Rows.Count;

            OleDbConnection DbConn = this.OpenXlsDbConnection(xlsFilePath);
            this.GetExcelDbAdapter(strSheetName);

            // Create a new sheet in the opening Excel file with specified name
            // -- Get column names for the new sheet
            string [] columnNames = new string[dtTable.Columns.Count];
            for (int i = 0; i < columnNames.Length; i++){
                columnNames[i] = dtTable.Columns[i].ColumnName;
            }
            // -- Generate a sql for create the new sheet
            string strCreateSql = GetCreateSheetSql(strSheetName, columnNames);

            // -- Execute the sql with OleDB 
            OleDbCommand cmd = new OleDbCommand(strCreateSql, DbConn);
            int effectRows = cmd.ExecuteNonQuery();

            FillSheetToDateSet(strSheetName);

            int rowIndex = 1;
            int colIndex = 0;

            //// Fill table column name to work-sheet
            //foreach (DataColumn col in dtTable.Columns)
            //{
            //    colIndex++;
            //    wrkSheet.Cells[rowIndex, colIndex] = col.ColumnName;
            //}

            //// Fill data-row to work-sheet
            //foreach (DataRow row in dtTable.Rows)
            //{
            //    rowIndex++;
            //    colIndex = 0;
            //    foreach (DataColumn col in dtTable.Columns)
            //    {
            //        colIndex++;
            //        wrkSheet.Cells[rowIndex, colIndex] = "'" + row[col.ColumnName].ToString();
            //    }
            //}

            //xlsApp.ActiveWorkbook.Save();
            //xlsApp.Quit();
            //xlsApp = null;
            GC.Collect();//垃圾回收

            return rowCount;
        }

        /// <summary>
        /// 方法：将DateSet 的指定 Table 保存到 Excel 文件中
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="FileName"></param>
        public void CreateExcel(DataSet ds, string tblName, string FileName)
        {
            throw (new Exception("This method is not implemented."));

            Page pg = new Page();
            HttpResponse resp;
            resp = pg.Response;
            resp.ContentEncoding = System.Text.Encoding.GetEncoding("GB2312");
            resp.AppendHeader("Content-Disposition", "attachment;filename=" + FileName);
            string colHeaders = "", ls_item = "";

            //定义表对象与行对象，同时用DataSet对其值进行初始化 
            DataTable dt = ds.Tables[tblName];
            DataRow[] myRow = dt.Select();//可以类似dt.Select("id>10")之形式达到数据筛选目的
            int i = 0;
            int cl = dt.Columns.Count;

            //取得数据表各列标题，各标题之间以t分割，最后一个列标题后加回车符 
            for (i = 0; i < cl; i++)
            {
                if (i == (cl - 1))//最后一列，加n
                {
                    colHeaders += dt.Columns[i].Caption.ToString() + "n";
                }
                else
                {
                    colHeaders += dt.Columns[i].Caption.ToString() + "t";
                }

            }
            resp.Write(colHeaders);
            //向HTTP输出流中写入取得的数据信息 

            //逐行处理数据   
            foreach (DataRow row in myRow)
            {
                //当前行数据写入HTTP输出流，并且置空ls_item以便下行数据     
                for (i = 0; i < cl; i++)
                {
                    if (i == (cl - 1))//最后一列，加n
                    {
                        ls_item += row[i].ToString() + "n";
                    }
                    else
                    {
                        ls_item += row[i].ToString() + "t";
                    }

                }
                resp.Write(ls_item);
                ls_item = "";

            }
            resp.End();
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
        private OleDbConnection OpenXlsDbConnection(string xlsFilePath)
        {
            if (this.xlsConnection == null)
            {
                // NOTE: Different version Excel file need specified connection string
                // string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + xlsFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                string strConnTmpl = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0 Xml;HDR=YES;IMEX=1;'";

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

                this.xlsDataSet.Dispose();
                this.xlsDataSet = null;
            }
        }

        // Fill excel sheet data to DateSet
        private int FillSheetToDateSet(string srcSheet = null)
        {
            if (srcSheet != null)
            {
                this.xlsSheetName = srcSheet;
            }
            this.xlsDataSet = new DataSet();

            // Fill data of the table in adapter to specified DataSet with given name
            return this.xlsDbAdapter.Fill(this.xlsDataSet, this.dataTableName);
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
            return this.xlsDbAdapter.Fill(trgDataSet, dataTableName);
        }

        public DataTable GetXlsDbTable()
        {
            return this.xlsDataSet.Tables[this.dataTableName];
        }

        public int AddToDbTable(DataAdapter dtAdapter, DataTable trgDbTable)
        {
            DataRow newRow = null;
            DataTable srcTable = GetXlsDbTable();

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

        public virtual void SaveToDb(DataSet srcDataSet = null)
        {

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
        }   */
    }
}
