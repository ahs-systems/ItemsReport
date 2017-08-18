using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace WindowsFormsApplication1
{
    public partial class frmReport : Form
    {
        public ItemsReport _parentForm;
        public bool firstLoad = true;
        public bool CloseTheForm = false;

        #region Disable Close Button
        //private const int CP_NOCLOSE_BUTTON = 0x200;
        //protected override CreateParams CreateParams
        //{
        //    get
        //    {
        //        CreateParams myCp = base.CreateParams;
        //        myCp.ClassStyle = myCp.ClassStyle | CP_NOCLOSE_BUTTON;
        //        return myCp;
        //    }
        //} 
        #endregion

        public frmReport()
        {
            InitializeComponent();
            cboYearPP.Items.Add(DateTime.Today.Year);
            cboYearPP.Items.Add(DateTime.Today.Year - 1);
            cboYearPP.SelectedIndex = 0;
        }

        private void frmReport_Load(object sender, EventArgs e)
        {
            //cboPP.SelectedIndex = 9;
            //cboItemsReport.SelectedIndex = 1;
        }

        public void LoadAllData()
        {
            if (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1)
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (firstLoad)
            {
                firstLoad = false;
                tabControl1.TabPages[1].Show();
            }

            Load_UUT_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());            
            Load_NPP_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_SC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_OC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_Terms_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_Trans_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
        }

        public void Load_Trans_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvTrans;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.UnitFrom, N.UnitTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM ItemsRpt_Transfers N JOIN SITES S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
                                 "N.PayPeriod_Year = @_PPYear and N.ItemsReportLetter = @_IRL ORDER BY N.Site, N.Emp_Name, N.EnteredDate";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        _dgv.DataSource = t;

                        foreach (DataGridViewColumn column in _dgv.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        //Hide Record ID Column
                        _dgv.Columns[0].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Load_Terms_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvTerms;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.TerminationDate, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM ItemsRpt_Terminations N JOIN SITES S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
                                 "N.PayPeriod_Year = @_PPYear and N.ItemsReportLetter = @_IRL ORDER BY N.Site, N.Emp_Name, N.EnteredDate";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        _dgv.DataSource = t;

                        foreach (DataGridViewColumn column in _dgv.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        //Hide Record ID Column
                        _dgv.Columns[0].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Load_UUT_Data(string _pp, string _ppYear, string _IRL)
        {
            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    dgvUUT.DataSource = null;
                    dgvUUT.Refresh();

                    string _sqlString = "Select S.SiteDesc, U.ID, U.Emp_Num, U.Emp_Name, U.UnitFrom, U.UnitTo, U.Occupation, " +
                                        "Case UPPER(U.ChangeInOccupation) " +
                                        "    When 'TRUE' Then NCHAR(0x394) " +
                                        "    Else '' " +
                                        "End as ' ', " +
                                        "U.Status, U.Comments, U.EnteredBy, U.EnteredDate, U.ChangeInSite " +
                                        "from ItemsRpt_UnitToUnitTransfer U join Sites S on U.Site = S.SiteID WHERE U.PayPeriod = @_PP AND " +
                                        "U.PayPeriod_Year = @_PPYear and U.ItemsReportLetter = @_IRL Order By U.Site, U.Emp_Name, U.EnteredDate";

                    

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        //da.SelectCommand.Parameters.AddWithValue("@S_NAME", "%" + _searchStr.ToUpper() + "%");
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        dgvUUT.DataSource = t;
                                               
                        foreach (DataGridViewColumn column in dgvUUT.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        foreach (DataGridViewRow row in dgvUUT.Rows)
                        {
                           if (row.Cells[dgvUUT.Columns.Count-1].Value.ToString().ToUpper() == "TRUE")
                            {
                                row.DefaultCellStyle.BackColor = Color.Yellow;                                
                            }
                        }                                                

                        //Hide Record ID and ChangeInSite Column
                        dgvUUT.Columns[1].Visible = dgvUUT.Columns[dgvUUT.Columns.Count-1].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Load_NPP_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvNPP;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();
                    
                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.Occupation, N.Status, N.EnteredBy, N.EnteredDate " +
                                 "FROM ItemsRpt_NewPrimaryPositions N JOIN SITES S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
                                 "N.PayPeriod_Year = @_PPYear and N.ItemsReportLetter = @_IRL ORDER BY N.Site, N.Emp_Name, N.EnteredDate";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        //da.SelectCommand.Parameters.AddWithValue("@S_NAME", "%" + _searchStr.ToUpper() + "%");
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        _dgv.DataSource = t;

                        foreach (DataGridViewColumn column in _dgv.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }                       

                        //Hide Record ID Column
                        _dgv.Columns[0].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Load_SC_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvSC;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();
                                        
                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.StatusFrom, N.StatusTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM ItemsRpt_StatusChange N JOIN SITES S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
                                 "N.PayPeriod_Year = @_PPYear and N.ItemsReportLetter = @_IRL ORDER BY N.Site, N.Emp_Name, N.EnteredDate";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {                        
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        _dgv.DataSource = t;

                        foreach (DataGridViewColumn column in _dgv.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        //Hide Record ID Column
                        _dgv.Columns[0].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Load_OC_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvOC;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.SystemsServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.OccFrom, N.OccTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                 "FROM ItemsRpt_OccupationChange N JOIN SITES S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
                                 "N.PayPeriod_Year = @_PPYear and N.ItemsReportLetter = @_IRL ORDER BY N.Site, N.Emp_Name, N.EnteredDate";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        da.SelectCommand.Parameters.AddWithValue("_PP", _pp);
                        da.SelectCommand.Parameters.AddWithValue("_PPYear", _ppYear);
                        da.SelectCommand.Parameters.AddWithValue("_IRL", _IRL);

                        DataTable t = new DataTable();
                        da.Fill(t);
                        _dgv.DataSource = t;

                        foreach (DataGridViewColumn column in _dgv.Columns)
                        {
                            column.SortMode = DataGridViewColumnSortMode.NotSortable;
                            column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        //Hide Record ID Column
                        _dgv.Columns[0].Visible = false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void dgvUUT_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvUUT, _parentForm.Load_UUT_Data);
        }

        private void btnEdit_UUT_Click(object sender, EventArgs e)
        {
            Edit(dgvUUT, _parentForm.Load_UUT_Data);
        }

        private void Edit(DataGridView _dgv, Action<string> _LoadMethod)
        {
            int _selectedRows = _dgv.SelectedCells.Cast<DataGridViewCell>()
                        .Select(cell => cell.OwningRow)
                        .Distinct()
                        .OrderBy(row => row.Index).ToArray().Length;

            if (_selectedRows > 1)
            {
                MessageBox.Show("You can only edit one row at a time.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (_dgv.CurrentCell == null)
            {
                MessageBox.Show("Please select first the record entry that you wan to edit.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            _parentForm.ID = _dgv.Rows[_dgv.CurrentCell.RowIndex].Cells["ID"].Value.ToString();
            _LoadMethod(_dgv.Rows[_dgv.CurrentCell.RowIndex].Cells["ID"].Value.ToString());
            Hide();
            _parentForm.Focus();
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {            
            LoadAllData();
        }

        private void btnDel_UUT_Click(object sender, EventArgs e)
        {
            //DeleteItem(dgvUUT.Rows[dgvUUT.CurrentCell.RowIndex].Cells["Emp_Name"].Value.ToString(), dgvUUT.Rows[dgvUUT.CurrentCell.RowIndex].Cells["ID"].Value.ToString(), "UUT");
        }

        private void DeleteItem(DataGridViewSelectedRowCollection _rows, string _tabName)
        {
            DialogResult _res;
            if (_rows.Count > 1)
            {
                _res = MessageBox.Show("Are you sure you want to delete the " + _rows.Count + " entries you have selected?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            else
            {
                _res = MessageBox.Show("Are you sure you want to delete the entry for: \n\n\"" + _rows[0].Cells["Emp_Name"].Value + "\" ?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            if (_res == DialogResult.Yes)
            {
                try
                {
                    using (SqlConnection _conn = new SqlConnection())
                    {
                        _conn.ConnectionString = Common.SystemsServer;
                        _conn.Open();

                        SqlCommand _command = _conn.CreateCommand();
                        foreach (DataGridViewRow _row in _rows)
                        {
                            _command.Parameters.Clear();

                            switch (_tabName)
                            {
                                case "NPP":
                                    _command.CommandText = "DELETE FROM ItemsRpt_NewPrimaryPositions WHERE Id = @_ID";
                                    break;
                                case "UUT":
                                    _command.CommandText = "DELETE FROM ItemsRpt_UnitToUnitTransfer WHERE Id = @_ID";
                                    break;
                                case "SC":
                                    _command.CommandText = "DELETE FROM ItemsRpt_StatusChange WHERE Id = @_ID";
                                    break;
                                case "OC":
                                    _command.CommandText = "DELETE FROM ItemsRpt_OccupationChange WHERE Id = @_ID";
                                    break;
                                case "Terms":
                                    _command.CommandText = "DELETE FROM ItemsRpt_Terminations WHERE Id = @_ID";
                                    break;
                                case "Trans":
                                    _command.CommandText = "DELETE FROM ItemsRpt_Transfers WHERE Id = @_ID";
                                    break;
                            }

                            _command.Parameters.AddWithValue("_ID", _row.Cells["ID"].Value);

                            _command.ExecuteNonQuery();
                        }

                        // Refresh the data
                        switch (_tabName)
                        {
                            case "NPP":
                                Load_NPP_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                            case "UUT":
                                Load_UUT_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                            case "SC":
                                Load_SC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                            case "OC":
                                Load_OC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                            case "Terms":
                                Load_Terms_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                            case "Trans":
                                Load_Trans_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
                                break;
                        }

                        MessageBox.Show("Successfully Deleted!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error in deleting the entry: " + ex.Message);
                }                        
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void dgvNPP_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvNPP, _parentForm.Load_NPP_Data);            
        }

        private void btnEdit_NPP_Click(object sender, EventArgs e)
        {
            Edit(dgvNPP, _parentForm.Load_NPP_Data);
        }

        private void btnDel_NPP_Click(object sender, EventArgs e)
        {
            var _dgv = (DataGridView) tabControl1.TabPages[tabControl1.SelectedIndex].Controls.Find(((Button)sender).Tag.ToString(), true).FirstOrDefault();            

            if (_dgv.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please highlight first the row(s) that you want to delete by clicking (and dragging if you want to select more than one row) the left most column.");
                return;
            }           
            //DeleteItem(_dgv.Rows[_dgv.CurrentCell.RowIndex].Cells["Emp_Name"].Value.ToString(), _dgv.Rows[_dgv.CurrentCell.RowIndex].Cells["ID"].Value.ToString(), _dgv.Parent.Name);
            DeleteItem(_dgv.SelectedRows, _dgv.Parent.Name);
        }

        private void btnEdit_SC_Click(object sender, EventArgs e)
        {
            Edit(dgvSC, _parentForm.Load_SC_Data);
        }

        private void btnEdit_OC_Click(object sender, EventArgs e)
        {
            Edit(dgvOC, _parentForm.Load_OC_Data);
        }

        private void btnEdit_Terms_Click(object sender, EventArgs e)
        {
            Edit(dgvTerms, _parentForm.Load_Terms_Data);
        }

        private void btnEdit_Trans_Click(object sender, EventArgs e)
        {
            Edit(dgvTrans, _parentForm.Load_Trans_Data);
        }

        private void dgvSC_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvSC, _parentForm.Load_SC_Data);
        }

        private void dgvOC_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvOC, _parentForm.Load_OC_Data);
        }

        private void dgvTerms_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvTerms, _parentForm.Load_Terms_Data);
        }

        private void dgvTrans_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            Edit(dgvTrans, _parentForm.Load_Trans_Data);
        }

        private void frmReport_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!CloseTheForm)
            {
                e.Cancel = true;
                //base.OnFormClosing(e);
                WindowState = FormWindowState.Minimized;
            }
        }
    }
}
