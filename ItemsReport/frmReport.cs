﻿using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ItemsReport
{
    public partial class frmReport : Form
    {
        public ItemsReport _parentForm;
        public bool firstLoad = true;
        public bool CloseTheForm = false;
        public bool LoadingNFP = false;

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
            cboYearPP.Items.Add(DateTime.Today.Year + 1);
            cboYearPP.Items.Add(DateTime.Today.Year);
            cboYearPP.Items.Add(DateTime.Today.Year - 1);
            cboYearPP.SelectedIndex = 1;
        }

        private void frmReport_Load(object sender, EventArgs e)
        {
            //cboPP.SelectedIndex = 9;
            //cboItemsReport.SelectedIndex = 1;
            dpNFPcheckingFrom.Value = DateTime.Today.AddDays(-7);
            dpNFPcheckingTo.Value = DateTime.Today;
            cboNFPchecking.SelectedIndex = 0;
        }

        public void LoadAllData()
        {
            if (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1)
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            // To properly show the highlighted rows on first load of the form
            if (firstLoad)
            {
                firstLoad = false;
                tabControl1.TabPages[1].Show();
                tabControl1.TabPages[6].Show();
            }

            Load_UUT_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_NPP_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_SC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_OC_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_Terms_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_Trans_Data(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString(), cboItemsReport.SelectedItem.ToString());
            Load_NFPChecking();
        }

        public void Load_Trans_Data(string _pp, string _ppYear, string _IRL)
        {
            DataGridView _dgv = dgvTrans;

            try
            {
                cboPP.SelectedItem = _pp; cboYearPP.SelectedItem = _ppYear; cboItemsReport.SelectedItem = _IRL;

                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.BooServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.UnitFrom, N.UnitTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM APP.ItemsRpt_Transfers N JOIN APP.Sites S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
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
                    _conn.ConnectionString = Common.BooServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.TerminationDate, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM APP.ItemsRpt_Terminations N JOIN APP.Sites S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
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
                    _conn.ConnectionString = Common.BooServer;

                    dgvUUT.DataSource = null;
                    dgvUUT.Refresh();

                    string _sqlString = "Select S.SiteDesc, U.ID, U.Emp_Num, U.Emp_Name, U.UnitFrom, U.UnitTo, U.Occupation, " +
                                        "Case UPPER(U.ChangeInOccupation) " +
                                        "    When 'TRUE' Then NCHAR(0x394) " +
                                        "    Else '' " +
                                        "End as ' ', " +
                                        "U.Status, U.Comments, U.EnteredBy, U.EnteredDate, U.ChangeInSite " +
                                        "from APP.ItemsRpt_UnitToUnitTransfer U join APP.Sites S on U.Site = S.SiteID WHERE U.PayPeriod = @_PP AND " +
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
                            // If change in site then highlight in yellow
                            if (row.Cells[dgvUUT.Columns.Count - 1].Value.ToString().ToUpper() == "TRUE")
                            {
                                row.DefaultCellStyle.BackColor = Color.Yellow;
                            }
                        }

                        //Hide Record ID and ChangeInSite Column
                        dgvUUT.Columns[1].Visible = dgvUUT.Columns[dgvUUT.Columns.Count - 1].Visible = false;
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
                    _conn.ConnectionString = Common.BooServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.Occupation, N.Status, N.EnteredBy, N.EnteredDate " +
                                 "FROM APP.ItemsRpt_NewPrimaryPositions N JOIN APP.Sites S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
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
                    _conn.ConnectionString = Common.BooServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.StatusFrom, N.StatusTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                "FROM APP.ItemsRpt_StatusChange N JOIN APP.Sites S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
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
                    _conn.ConnectionString = Common.BooServer;

                    _dgv.DataSource = null;
                    _dgv.Refresh();

                    string _sqlString = "SELECT ID, S.SiteDesc, N.Emp_Num, N.Emp_Name, N.Unit, N.OccFrom, N.OccTo, N.Comments, N.EnteredBy, N.EnteredDate " +
                                 "FROM APP.ItemsRpt_OccupationChange N JOIN APP.Sites S ON N.Site = S.SiteID WHERE N.PayPeriod = @_PP AND " +
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

        public DataTable GetNFPList()
        {
            DataTable _ret = new DataTable();

            try
            {
                using (SqlConnection _conn = new SqlConnection())
                {
                    _conn.ConnectionString = Common.BooServer;

                    string _filter = "";
                    switch (cboNFPchecking.SelectedIndex)
                    {
                        case 0:
                            _filter = "";
                            break;
                        case 1:
                            _filter = " AND UPPER(Prev_PayInfo) LIKE '%NOT FOR PAYROLL%'";
                            break;
                        case 2:
                            _filter = " AND UPPER(Prev_PayInfo) LIKE '%INACTIVE%'";
                            break;
                    }

                    string _sqlString = "SELECT * FROM APP.NFPChecking WHERE (CurrentStat = 0 OR DateUploaded BETWEEN @_from and @_to)" + _filter + " ORDER BY DateUploaded";

                    using (SqlDataAdapter da = new SqlDataAdapter(_sqlString, _conn))
                    {
                        da.SelectCommand.Parameters.AddWithValue("_from", dpNFPcheckingFrom.Value.ToString("dd-MMM-yyyy"));
                        da.SelectCommand.Parameters.AddWithValue("_to", dpNFPcheckingTo.Value.AddDays(1).ToString("dd-MMM-yyyy"));
                        da.Fill(_ret);
                    }

                    return _ret;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        public void Load_NFPChecking()
        {
            DataGridView _dgv = dgvNFPChecking;

            try
            {
                LoadingNFP = true;

                _dgv.DataSource = null;

                _dgv.DataSource = GetNFPList();

                if (_dgv.DataSource != null)
                {

                    foreach (DataGridViewColumn column in _dgv.Columns)
                    {
                        column.SortMode = DataGridViewColumnSortMode.NotSortable;
                        // Fill the length of the grid with the last column
                        //if (column.Name != "Comments")
                        //    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        //else
                        //    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                        column.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                    }

                    //foreach (DataGridViewColumn column in _dgv.Columns)
                    //{
                    //    int colw = column.Width;
                    //    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    //    column.Width = colw;
                    //}

                    foreach (DataGridViewColumn _col in _dgv.Columns)
                    {
                        if (_col.Name != "CurrentStat" && _col.Name != "Comments")
                        {
                            _col.ReadOnly = true;
                        }
                    }

                    //Hide Record ID Column
                    _dgv.Columns[0].Visible = false;

                    foreach (DataGridViewRow row in _dgv.Rows)
                    {
                        // If not yet check then highlight them
                        if (row.Cells["CurrentStat"].Value.ToString().ToUpper() == "FALSE")
                        {
                            row.DefaultCellStyle.BackColor = Color.Orange;
                        }
                    }


                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                LoadingNFP = false;
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
                        _conn.ConnectionString = Common.BooServer;
                        _conn.Open();

                        SqlCommand _command = _conn.CreateCommand();
                        foreach (DataGridViewRow _row in _rows)
                        {
                            _command.Parameters.Clear();

                            switch (_tabName)
                            {
                                case "NPP":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_NewPrimaryPositions WHERE Id = @_ID";
                                    break;
                                case "UUT":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_UnitToUnitTransfer WHERE Id = @_ID";
                                    break;
                                case "SC":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_StatusChange WHERE Id = @_ID";
                                    break;
                                case "OC":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_OccupationChange WHERE Id = @_ID";
                                    break;
                                case "Terms":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_Terminations WHERE Id = @_ID";
                                    break;
                                case "Trans":
                                    _command.CommandText = "DELETE FROM APP.ItemsRpt_Transfers WHERE Id = @_ID";
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
            var _dgv = (DataGridView)tabControl1.TabPages[tabControl1.SelectedIndex].Controls.Find(((Button)sender).Tag.ToString(), true).FirstOrDefault();

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

        private void dgvNFPChecking_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (LoadingNFP)
            {
                return;
            }

            try
            {
                using (SqlConnection _conn = new SqlConnection(Common.BooServer))
                {
                    _conn.Open();
                    using (SqlCommand _comm = _conn.CreateCommand())
                    {
                        _comm.CommandText = "UPDATE APP.NFPChecking SET CheckedBy = @_currUser,  CheckedDate = getdate(), CurrentStat = @_stat, Comments = @_comments  WHERE ID = @_id";
                        _comm.Parameters.AddWithValue("_currUser", Common.CurrentUser);
                        if (dgvNFPChecking.CurrentRow.Cells["Comments"].Value.ToString() != "")
                        {
                            _comm.Parameters.AddWithValue("_stat", "1");
                        }
                        else
                        {
                            _comm.Parameters.AddWithValue("_stat", (bool)dgvNFPChecking.CurrentRow.Cells["CurrentStat"].Value ? "1" : "0");
                        }
                        _comm.Parameters.AddWithValue("_id", dgvNFPChecking.CurrentRow.Cells["id"].Value.ToString());
                        _comm.Parameters.AddWithValue("_comments", dgvNFPChecking.CurrentRow.Cells["Comments"].Value.ToString());
                        _comm.ExecuteNonQuery();
                    }
                }

                int _currentRowIndex = dgvNFPChecking.CurrentRow.Index;
                //MessageBox.Show("Changes was successfuly saved.", "Confirm", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Load_NFPChecking();

                dgvNFPChecking.FirstDisplayedScrollingRowIndex = _currentRowIndex;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in NFP Checking: " + ex.Message);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Load_NFPChecking();
        }

        private void btnNFPtoExcel_Click(object sender, EventArgs e)
        {

            DataTable t = GetNFPList();

            if (t == null)
            {
                MessageBox.Show("Error in getting NFP and Inactive list", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                using (var package = new ExcelPackage())
                {
                    // add a new worksheet to the empty workbook
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("From NFP and Inactive");

                    // Set Page Settings
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    worksheet.PrinterSettings.ShowGridLines = true;
                    worksheet.PrinterSettings.HorizontalCentered = true;
                    worksheet.PrinterSettings.TopMargin = (decimal)1.5 / 2.54M;
                    worksheet.PrinterSettings.BottomMargin = (decimal)1.5 / 2.54M;
                    worksheet.PrinterSettings.LeftMargin = (decimal)0.25 / 2.54M;
                    worksheet.PrinterSettings.RightMargin = (decimal)0.25 / 2.54M;
                    worksheet.PrinterSettings.HeaderMargin = (decimal)0.5 / 2.54M;
                    worksheet.PrinterSettings.FooterMargin = (decimal)0.5 / 2.54M;
                    worksheet.HeaderFooter.OddHeader.LeftAlignedText = DateTime.Now.ToString("ddMMMyyyy HH:mm:ss");
                    worksheet.HeaderFooter.OddHeader.RightAlignedText = "";
                    worksheet.HeaderFooter.OddHeader.CenteredText = "From NFP and Inactive to ESP";
                    worksheet.HeaderFooter.OddFooter.RightAlignedText = string.Format("Page {0} of {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                    worksheet.View.PageBreakView = true;
                    worksheet.PrinterSettings.RepeatRows = new ExcelAddress("$1:$1");
                    worksheet.PrinterSettings.FitToPage = true; worksheet.PrinterSettings.FitToWidth = 1; worksheet.PrinterSettings.FitToHeight = 0;

                    //Setting Header Style
                    worksheet.Row(1).Height = 25;
                    worksheet.Cells[1, 1].Value = "Record Type"; //worksheet.Column(2).Width = 12.30;
                    worksheet.Cells[1, 2].Value = "Date Uploaded"; //worksheet.Column(3).Width = 10.43; //worksheet.Column(3).AutoFit(); //
                    worksheet.Cells[1, 3].Value = "Emp ID"; //worksheet.Column(4).Width = 22;
                    worksheet.Cells[1, 4].Value = "Name"; //worksheet.Column(5).Width = 35;
                    worksheet.Cells[1, 5].Value = "Previous Unit"; //worksheet.Column(6).Width = 35;
                    worksheet.Cells[1, 6].Value = "Comments"; //worksheet.Column(7).Width = 9.86; worksheet.Cells[1, 7].Style.WrapText = true;

                    var range = worksheet.Cells[1, 1, 1, 6];
                    range.Style.Font.Bold = true;
                    range.Style.Font.Size = 11;
                    range.Style.Font.Name = "Verdana";


                    int lineCtr = 2;

                    foreach (DataRow _row in t.Rows)
                    {
                        if (_row["CheckedBy"].ToString().Trim() == "")
                        {
                            worksheet.Row(lineCtr).Height = 25;
                            worksheet.Row(lineCtr).Style.Font.Name = "Verdana";
                            worksheet.Row(lineCtr).Style.Font.Size = 11;
                            worksheet.Row(lineCtr).Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            worksheet.Cells[lineCtr, 1].Value = _row["Type"];
                            worksheet.Cells[lineCtr, 2].Value = Convert.ToDateTime(_row["DateUploaded"]).ToString("dd-MMM-yyyy HH:mm");
                            worksheet.Cells[lineCtr, 3].Value = _row["EmpID"];
                            worksheet.Cells[lineCtr, 4].Value = _row["Name"];
                            worksheet.Cells[lineCtr, 5].Value = _row["Prev_PayInfo"];
                            worksheet.Cells[lineCtr, 6].Value = _row["Comments"];
                            lineCtr++;

                            if (lineCtr % 9 == 0) worksheet.Row(lineCtr).PageBreak = true;
                        }
                    }

                    worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                    worksheet.Column(6).Width = 38; // expand the width for column "Comments"                                       

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 1;
                    saveFileDialog1.FileName = DateTime.Today.ToString("dd-MMM-yyyy ") + "From NFP and Inactive";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        package.SaveAs(new FileInfo(saveFileDialog1.FileName));
                        System.Diagnostics.Process.Start(saveFileDialog1.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnRunCheck_Click(object sender, EventArgs e)
        {
            try
            {
                btnRunCheck.Text = "Please wait...";
                btnRunCheck.Update();

                using (SqlConnection _conn = new SqlConnection(Common.BooServer))
                {
                    SqlCommand _comm = _conn.CreateCommand();
                    _comm.CommandText = "select * from APP.NFPChecking where CurrentStat = 0";

                    _conn.Open();
                    SqlDataReader _dr = _comm.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        while (_dr.Read())
                        {
                            string _payInfo = GetPayInfo(_dr["EmpID"].ToString());

                            // if _payInfo == "" then something went wrong, don't do anything, otherwise proceed
                            if (_payInfo != "")
                            {
                                UpdateNFPcheckingList(_dr["id"].ToString(), _payInfo, _dr["EmpID"].ToString(), _dr["type"].ToString());
                            }
                        }

                        Load_NFPChecking();
                    }

                    _conn.Close();

                    MessageBox.Show("Done checking.", "Result", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in Run Checking: " + ex.Message);
            }
            finally
            {
                btnRunCheck.Text = "Run Auto Checking";
            }

        }

        private void UpdateNFPcheckingList(string _id, string _payInfo, string _empID, string _fileType)
        {
            try
            {
                string _comment = "";

                // if Pay Info is no longer Inactive or NFP
                if (!"Not for Payroll , --- INACTIVE ---, Inactive".ToUpper().Contains(_payInfo.ToUpper()))
                {
                    _comment = GetWhoChangedPayInfo(_empID);
                }
                // if Pay Info is still Inactive or NFP then don't update the Current Pay Info, get who last changed the Positions tab
                else
                {
                    _payInfo = "";
                    _comment = GetWhoChangePositions(_empID);
                }

                // if the last update was not made by RSSS then change the comment
                if (!_comment.Contains("RSSS"))
                {
                    _comment = "*** This came from Record Type " + _fileType;
                }

                using (SqlConnection _conn = new SqlConnection(Common.BooServer))
                {
                    _conn.Open();
                    using (SqlCommand _comm = _conn.CreateCommand())
                    {
                        _comm.CommandText = "UPDATE APP.NFPChecking SET CheckedBy = 'AutoSystem',  CheckedDate = getdate(), Curr_PayInfo = @_currUnit, Comments = @_comments  WHERE ID = @_id";
                        _comm.Parameters.AddWithValue("_currUnit", _payInfo);
                        _comm.Parameters.AddWithValue("_comments", _comment);
                        _comm.Parameters.AddWithValue("_id", _id);
                        _comm.ExecuteNonQuery();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in UpdateNFPcheckingList: " + ex.Message);
            }
        }

        private string GetWhoChangePositions(string _empID)
        {
            string _ret = "";

            try
            {
                using (SqlConnection _conn = new SqlConnection(Common.ESPServer))
                {
                    _conn.Open();
                    using (SqlCommand _comm = _conn.CreateCommand())
                    {
                        _comm.CommandText = "Select top 1 Format(EP.EP_ChangeDate,'dd-MMM-yyyy hh:mm tt') AS ChangeDate, Users.US_FullName from EmpPosition EP " +
                                "left join Users on ep.EP_ChangeUserID = Users.US_UserID " +
                                "where EP.EP_EmpID in (select DISTINCT E_EmpID from emp where E_EmpNbr LIKE @_empID) " +
                                "AND EP_ToDate > GETDATE() ORDER BY EP_ChangeDate DESC";

                        _comm.Parameters.AddWithValue("_empID", _empID.Substring(0, 8) + "%");

                        SqlDataReader _dr = _comm.ExecuteReader();

                        if (_dr.HasRows)
                        {
                            _dr.Read();
                            _ret = "Position last changed by: " + _dr["US_FullName"].ToString().Trim() + " on " + _dr["ChangeDate"];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in GetWhoChangedPayInfo: " + ex.Message);
            }

            return _ret;
        }

        private string GetWhoChangedPayInfo(string _empID)
        {
            string _ret = "";

            try
            {
                using (SqlConnection _conn = new SqlConnection(Common.ESPServer))
                {
                    _conn.Open();
                    using (SqlCommand _comm = _conn.CreateCommand())
                    {
                        _comm.CommandText = "select TOP 1 Format(ETCI_ChangeDate,'dd-MMM-yyyy hh:mm tt') AS [ChangeDate], (select US_FullName from Users where US_UserID = ETCI.ETCI_ChangeUserID) as ChangeUser from EmpTimeCardInfo ETCI where ETCI_EmpID = " +
                            "(select E_EmpID from emp where E_EmpNbr = @EMPID) " +
                            //"AND ETCI_PayPeriodID <= (select PP_PayPeriodID from PayPeriod where getdate() between PP_StartDate and PP_EndDate) " +
                            "ORDER BY ETCI_PayPeriodID DESC";

                        _comm.Parameters.AddWithValue("EMPID", _empID);

                        SqlDataReader _dr = _comm.ExecuteReader();

                        if (_dr.HasRows)
                        {
                            _dr.Read();
                            _ret = "PayInfo changed by: " + _dr["ChangeUser"].ToString().Trim() + " on " + _dr["ChangeDate"];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in GetWhoChangedPayInfo: " + ex.Message);
            }

            return _ret;

        }

        private string GetPayInfo(string _empNbr)
        {
            string _ret = "";

            try
            {
                using (SqlConnection _conn = new SqlConnection(Common.ESPServer))
                {
                    SqlCommand _comm = _conn.CreateCommand();

                    _comm.CommandText = "SELECT TCG_Desc from TimeCardGroup where TCG_TCardGroupID = " +
                                        "(select TOP 1 ETCI_TimeCardGroupID from EmpTimeCardInfo ETCI where ETCI_EmpID = " +
                                        "(select E_EmpID from emp where E_EmpNbr = @_empID) ORDER BY ETCI_PayPeriodID DESC)";

                    //_comm.CommandText = "SELECT TCG_Desc from TimeCardGroup where TCG_TCardGroupID =  " +
                    //                    "(select TOP 1 ETCI_TimeCardGroupID from EmpTimeCardInfo ETCI where ETCI_EmpID = " +
                    //                    "(select E_EmpID from emp where E_EmpNbr = @_empID) " +
                    //                    "AND ETCI_PayPeriodID <= (select PP_PayPeriodID from PayPeriod where getdate() between PP_StartDate and PP_EndDate) " +
                    //                    "ORDER BY ETCI_PayPeriodID DESC)";

                    _comm.Parameters.AddWithValue("_empID", _empNbr);

                    _conn.Open();

                    SqlDataReader _dr = _comm.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        _ret = _dr["TCG_Desc"].ToString().Trim();
                    }
                    else
                    {
                        _ret = "--- INACTIVE ---";
                    }

                    _conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in GetPayInfo: " + ex.Message);
            }

            return _ret;
        }


    }
}
