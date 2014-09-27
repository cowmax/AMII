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

        int AddToDb()
        {
            return 0;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            ExcelDb etb = new ExcelDb("..\\..\\xls\\西伍商业决策分析系统_功能框架列表.xlsx", "系统用户权限分类", "srcTable");

            DataTable srcTable = etb.GetXlsDbTable();

            dataGridView1.DataSource = srcTable;

            // Save to DB's table
            SystemUserPrivilegeTableAdapter pvlgTAdapter = new SystemUserPrivilegeTableAdapter();
            GwmsTestDataSet dtSet = new GwmsTestDataSet();

            GwmsTestDataSet.SystemUserPrivilegeDataTable pvlgTable = new GwmsTestDataSet.SystemUserPrivilegeDataTable();
            GwmsTestDataSet.SystemUserPrivilegeRow pvlgRow = null;

            foreach (DataRow r in srcTable.Rows)
            {
                pvlgRow = pvlgTable.NewSystemUserPrivilegeRow();

                // set each field value
                foreach (DataColumn col in pvlgTable.Columns)
                {
                    pvlgRow[col.Caption] = r[col.Caption];
                }

                pvlgTable.AddSystemUserPrivilegeRow(pvlgRow);
            }

            pvlgTAdapter.Update(pvlgTable);

            etb.Close();
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            pmmaTableAdapter pmmaTblAdapter = new pmmaTableAdapter();
            GwmsTestDataSet wmsDataSet = new GwmsTestDataSet();
            GwmsTestDataSet.pmmaDataTable pmmaTable = new GwmsTestDataSet.pmmaDataTable();

            int rowCount = pmmaTblAdapter.Fill(pmmaTable);

            ExcelDb etb = new ExcelDb(pmmaTable);

            etb.SaveToXls("D:\\百度云\\Project\\ImportExcelToDB\\ImportExcelToDB\\xls\\西伍商业决策分析系统_功能框架列表.xlsx", "dataTable");
        }
    }

    class PvlgTable : GwmsTestDataSet.SystemUserPrivilegeDataTable
    {
        DataRowCollection getRows()
        {
            return this.Rows;
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

            return this.Adapter.Update(this._table);
        }
    }
}
