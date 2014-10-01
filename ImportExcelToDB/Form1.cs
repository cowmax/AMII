using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ImportExcelToDB.GwmsTestDataSetTableAdapters;
using System.Data.SqlClient;

namespace ImportExcelToDB
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // 1. 初始化 ExcelDb 实例：指定Excel 文件路径、Sheet 的名称
            ExcelDb etb = new ExcelDb("..\\..\\xls\\西伍商业决策分析系统_功能框架列表.xls", "系统用户权限分类");

            // 2. 初始化完成后，ExcelDb 已经将 Excel 文件 的 Sheet 中的数据加载到 DataTable 里
            DataTable srcTable = etb.GetOleDbTable();

            dataGridView1.DataSource = srcTable; // 在 DataGridView 中显示 DataTable 数据

            // 3. 把 DataTable 数据保存到数据库中
            SystemUserPrivilegeTableAdapter trgDbAdapter = new SystemUserPrivilegeTableAdapter();
            GwmsTestDataSet.SystemUserPrivilegeDataTable trgTable = new GwmsTestDataSet.SystemUserPrivilegeDataTable();

            // -- 逐行逐列复制数据
            foreach (DataRow srcRow in srcTable.Rows)
            {
                DataRow tmpRow = trgTable.NewRow();
                foreach (DataColumn col in trgTable.Columns)
                {
                    tmpRow[col.ColumnName] = srcRow[col.ColumnName];
                }
                trgTable.Rows.Add(tmpRow);
            }

            // -- 提交到数据库
            int eftRow = trgDbAdapter.Update(trgTable);

            etb.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pmmaTableAdapter pmmaTblAdapter = new pmmaTableAdapter();
            GwmsTestDataSet wmsDataSet = new GwmsTestDataSet();
            GwmsTestDataSet.pmmaDataTable pmmaTable = new GwmsTestDataSet.pmmaDataTable();

            int rowCount = pmmaTblAdapter.Fill(pmmaTable);

            string xlsFilePath = "D:\\百度云\\Project\\ImportExcelToDB\\ImportExcelToDB\\xls\\西伍商业决策分析系统_功能框架列表.xls";

            ExcelDb etb = new ExcelDb(pmmaTable);

            etb.SaveToXlsFile(xlsFilePath, "Test");

            // etb.SaveToXlsFile(xlsFilePath, "dataTable");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            pmmaTableAdapter pmmaTblAdapter = new pmmaTableAdapter();
            GwmsTestDataSet wmsDataSet = new GwmsTestDataSet();
            GwmsTestDataSet.pmmaDataTable pmmaTable = new GwmsTestDataSet.pmmaDataTable();

            int rowCount = pmmaTblAdapter.Fill(pmmaTable);

            ExcelDb etb = new ExcelDb(pmmaTable);

            etb.SaveToXls("D:\\百度云\\Project\\ImportExcelToDB\\ImportExcelToDB\\xls\\西伍商业决策分析系统_功能框架列表.xls", "dataTable");
        }

        private void button7_Click(object sender, EventArgs e)
        {


            
        }

        private void button8_Click(object sender, EventArgs e)
        {
            string templateFolder = "..\\..\\Templates";
            TemplateManager tmplMgr = new TemplateManager(templateFolder);

        }
    }

    class PvlgTable : GwmsTestDataSet.SystemUserPrivilegeDataTable
    {
        DataRowCollection getRows()
        {
            return base.Rows;
        }
    }

    class PvlgTabAdapter : SystemUserPrivilegeTableAdapter
    {
        PvlgTable _table = null;

        public PvlgTabAdapter(PvlgTable tbl)
        {
            this._table = tbl;
        }

        public int AddToDbTable(DataTable srcTable)
        {
            DataRow newRow = null;

            foreach (DataRow srcRow in srcTable.Rows)
            {
                newRow = this._table.NewRow();

                // set each field value
                foreach (DataColumn col in this._table.Columns)
                {
                    newRow[col.Caption] = srcRow[col.Caption];
                }

                this._table.Rows.Add(newRow);
            }

            return base.Adapter.Update(this._table);
        }
    }
}
