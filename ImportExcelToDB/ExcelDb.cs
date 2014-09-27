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
        string xlsFilePath;
        string strSheetName;
        OleDbConnection xlsConnection;
        OleDbDataAdapter xlsDbAdapter;
        DataSet xlsDataSet = new DataSet(); // 用于保存 Excel/Sheet 数据的 DataSet 对像
        string xlsTableName;

        public ExcelDb(DataTable dtTable)
        {
            this.xlsDataSet.Tables.Add(dtTable);
            this.xlsTableName = dtTable.TableName;
            this.strSheetName = dtTable.TableName;
        }

        /// <summary>
        /// 构造函数：从Excel文件的指定Sheet 中加载数据到 DataSet 中
        /// </summary>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中的 Sheet 名称</param>
        /// <param name="srcTableName">数据库集中相应的 Table名称，缺省名称与Sheet 相同</param>
        public ExcelDb(string xlsFilePath, string strSheetName, string srcTableName = null)
        {
            LoadFromXls(xlsFilePath, strSheetName, srcTableName);
        }

        /// <summary>
        /// 构造函数：从Excel文件的指定Sheet 中加载数据到 DataSet 中
        /// </summary>
        /// <param name="xlsFilePath">Excel 文件路径</param>
        /// <param name="strSheetName">Excel 中的 Sheet 名称</param>
        /// <param name="srcTableName">数据库集中相应的 Table名称，缺省名称与Sheet 相同</param>
        public int LoadFromXls(string xlsFilePath, string strSheetName, string srcTableName = null)
        {
            this.xlsFilePath = xlsFilePath;
            this.strSheetName = strSheetName;

            if (srcTableName == null)
            {
                this.xlsTableName = GetTableName(strSheetName);
            }
            else
            {
                this.xlsTableName = srcTableName;
            }

            this.Open();
            this.GetExcelDbAdapter();
            return this.Fill();
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
            DataTable dtTable = this.xlsDataSet.Tables[this.xlsTableName];
            int rowCount = dtTable.Rows.Count;

            // NOTE : Must Add reference Microsoft.Office.Interop.Excel assembly
            Excel.Application xlsApp = new Excel.Application();
            Excel.Workbook xlsWrkBook= xlsApp.Application.Workbooks.Open(xlsFilePath);
            Excel.Worksheet wrkSheet = getSheet(xlsWrkBook.Sheets, this.xlsTableName, true);

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

            xlsApp.ActiveWorkbook.Save();
            xlsApp.Quit();
            xlsApp = null;
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

        private string GetTableName(string sheetName){
            return sheetName;
        }

        private OleDbConnection Open()
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.xlsFilePath + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
            //连接数据源      
            this.xlsConnection = new OleDbConnection(strConn);
            this.xlsConnection.Open();

            return this.xlsConnection;
        }

        private void CloseExcelFileConnection()
        {
            if (this.xlsConnection != null)
            {
                this.xlsConnection.Close();
            }
        }

        public void Close()
        {
            CloseExcelFileConnection();
        }

        private int Fill()
        {
           return this.xlsDbAdapter.Fill(this.xlsDataSet, this.xlsTableName);
        }

        private string GetSelectAllSql()
        {
            return ("select * from  [" + this.strSheetName + "$] ");
        }

        private OleDbDataAdapter GetExcelDbAdapter()
        {
            // 打开数据源
            OleDbConnection conn = Open();

            string strSql = GetSelectAllSql();
            this.xlsDbAdapter = new OleDbDataAdapter(strSql, conn);

            return this.xlsDbAdapter;
        }

        private int Fill(DataSet trgDataSet, string srcTable)
        {
            return this.xlsDbAdapter.Fill(trgDataSet, xlsTableName);
        }

        public DataTable GetXlsDbTable()
        {
            return this.xlsDataSet.Tables[this.xlsTableName];
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
