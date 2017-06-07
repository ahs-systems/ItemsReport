using System;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml;
using System.IO;

namespace WindowsFormsApplication1
{
    public partial class ItemsReport : Form
    {
        AutoCompleteStringCollection unitsShortDesc = new AutoCompleteStringCollection();
        AutoCompleteStringCollection unitsLongDesc = new AutoCompleteStringCollection();

        frmReport _frmReport = new frmReport();

        public string pp;
        public string ppYear;
        public string itemsReportLetter;
        public string ID;


        public ItemsReport()
        {
            InitializeComponent();
        }

        private void ItemsReport_Load(object sender, EventArgs e)
        {
            Common.CurrentUser = System.Security.Principal.WindowsIdentity.GetCurrent().Name;

            _frmReport._parentForm = this;
            
            cboYearPP.Items.Add(DateTime.Today.Year); cboYearPP.Items.Add(DateTime.Today.Year - 1);
            cboYearPP.SelectedIndex = 0;            

            cboPP.SelectedItem = Common.GetPP(DateTime.Now.ToString("yyyy-MM-dd"));

            PopulateUnitShortDesc(ref unitsShortDesc);                       
            txtUnit_NPP.AutoCompleteCustomSource = txtTransFrom_UUT.AutoCompleteCustomSource = txtTransTo_UUT.AutoCompleteCustomSource = txtUnit_SC.AutoCompleteCustomSource = txtUnit_OC.AutoCompleteCustomSource = unitsShortDesc;

            PopulateUnitLongDesc(ref unitsLongDesc);
            txtUnit_Terms.AutoCompleteCustomSource = txtUnitFrom_Trans.AutoCompleteCustomSource = unitsLongDesc;

            //cboPP.SelectedIndex = 9;
            //cboItemsReport.SelectedIndex = 1;
        }

        private int GetSiteNum_ShortDesc(string _unitShortDesc)
        {
            int _ret = -1;
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "select Substring(U_Desc,1,2) AS U_PREFIX from unit where UPPER(U_ShortDesc) = UPPER(@_ShortDesc)";
                    myCommand.Parameters.AddWithValue("_ShortDesc", _unitShortDesc);

                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        myReader.Read();
                        if (myReader["U_PREFIX"].ToString() == "S2")
                        {
                            _ret = 5;
                        }
                        else
                        {
                            _ret = Convert.ToInt16(myReader["U_PREFIX"]) - 1;
                        }
                        if (_ret > 5) // invalid site number
                        {
                            _ret = -1;
                        }
                    }
                    myCommand.Dispose();
                }               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error (GetSiteNum_ShortDesc): " + ex.Message, "ERROR");
            }

            return _ret;
        }

        private int GetSiteNum_LongDesc(string _unitLongDesc)
        {
            int _ret = -1;
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "select Substring(U_Desc,1,2) AS U_PREFIX from unit where UPPER(U_Desc) = UPPER(@_LongDesc)";
                    myCommand.Parameters.AddWithValue("_LongDesc", _unitLongDesc);

                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        myReader.Read();
                        if (myReader["U_PREFIX"].ToString() == "S2")
                        {
                            _ret = 5;
                        }
                        else
                        {
                            _ret = Convert.ToInt16(myReader["U_PREFIX"]) - 1;
                        }
                        if (_ret > 5) // invalid site number
                        {
                            _ret = -1;
                        }
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error (GetSiteNum_LongDesc): " + ex.Message, "ERROR");
            }

            return _ret;
        }

        private void PopulateUnitLongDesc(ref AutoCompleteStringCollection _unitsLongDesc)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "select U_Desc from unit where (U_Desc like '0%' OR U_Desc like 'S%') AND U_Active = 1 ORDER BY U_DESC";

                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        while (myReader.Read())
                            _unitsLongDesc.Add(myReader["U_Desc"].ToString().Trim());
                    }

                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void PopulateUnitShortDesc(ref AutoCompleteStringCollection _unitSource)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "select U_ShortDesc from unit where (U_Desc like '0%' OR U_Desc like 'S%') AND U_Active = 1 ORDER BY U_ShortDesc";
                                        
                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        while(myReader.Read())
                        _unitSource.Add(myReader["U_ShortDesc"].ToString().Trim());
                    }

                    myCommand.Dispose();                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");                
            }
        }

        private void btnSave_NPP_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_NPP.SelectedIndex == -1 || txtEmpNo_NPP.Text.Trim() == "" || txtEmpName_NPP.Text.Trim() == "" || txtUnit_NPP.Text.Trim() == "" || txtOcc_NPP.Text.Trim() == "" || txtStatus_NPP.Text.Trim() == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_NPP.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_NPP.Focus();
                return;
            }

            try
            {                

                using (SqlConnection myConnection = new SqlConnection())
                {
                    //myConnection.ConnectionString = @"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (((Button)sender).Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_NewPrimaryPositions (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, Unit, Occupation, Status, EnteredBy) values " +
                        "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_Unit, @_Occupation, @_Status, @_EnteredBy)";
                    }
                    else // Update
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "Update ItemsRpt_NewPrimaryPositions SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, " +
                            "PayPeriod_Year = @_PayPeriod_Year, Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, Unit = @_Unit, Occupation = @_Occupation, " +
                            "Status = @_Status, EnteredBy = @_EnteredBy WHERE ID = " + ID;                       
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_NPP.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_NPP.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_NPP.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Unit", txtUnit_NPP.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Occupation", txtOcc_NPP.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Status", txtStatus_NPP.Text.Trim().ToUpper());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 0, "NPP");                        
                    }

                    _frmReport.Load_NPP_Data(_pp, _ppYear, _ItemsReportLetter);
                    _frmReport.Show();
                    _frmReport.tabControl1.SelectedIndex = 0;
                    ClearForm(tabControl1.TabPages[0]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void txtEmpNo_NPP_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Trim().Length > 7)
            {
                string _ret = SearchEmpName(((TextBox)sender).Text.Substring(0,8));
                var _empNameTextBox = this.Controls.Find(((TextBox)sender).Tag.ToString(), true).FirstOrDefault();
                ((TextBox)_empNameTextBox).Text = _ret;                
            }            
        }

        private string SearchEmpName(string _empNo)
        {
            string _ret = "";

            try
            {

                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "SELECT LTRIM(RTRIM(E_LASTNAME)) + ', ' + LTRIM(RTRIM(E_FIRSTNAME)) 'DESC' FROM EMP WHERE E_EMPNBR LIKE @V_SEARCH AND LEN(E_EMPNBR) > 7";

                    myCommand.Parameters.Add(new SqlParameter("V_SEARCH", _empNo + "%"));
                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        myReader.Read();
                        _ret = myReader["DESC"].ToString();
                    }                   

                    myCommand.Dispose();

                    return _ret;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
                return _ret;
            }
        }

        private void txtOccCode_NPP_TextChanged(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Trim().Length > 2)
            {
                string _ret = SearchOccupation(((TextBox)sender).Text.Trim());
                var _empNameTextBox = this.Controls.Find(((TextBox)sender).Tag.ToString(), true).FirstOrDefault();
                ((TextBox)_empNameTextBox).Text = _ret;
            }            
        }

        private string SearchOccupation(string _code)
        {
            string _ret = "";

            try
            {

                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.ESPServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "SELECT LTRIM(RTRIM(O_CODE)) + ' - ' + O_DESC 'DESC' FROM OCCUPATION WHERE O_CODE LIKE @V_O_CODE order by o_code";

                    myCommand.Parameters.Add(new SqlParameter("V_O_CODE", _code + "%"));
                    SqlDataReader myReader = myCommand.ExecuteReader();

                    if (myReader.HasRows)
                    {
                        myReader.Read();
                        _ret = myReader["DESC"].ToString();
                    }

                    myCommand.Dispose();

                    return _ret;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
                return _ret;
            }
        }

        private string GetStartPP(string _pp, string _ppYear)
        {
            string _ret = "";

            using (SqlConnection myConnection = new SqlConnection())
            {
                myConnection.ConnectionString = Common.ESPServer;
                myConnection.Open();

                SqlCommand _comm = myConnection.CreateCommand();
                _comm.CommandText = "select Format(PP_StartDate,'MMM dd, yyyy') AS StartDate From payperiod where PP_Nbr = @_PP and Year(PP_StartDate) = @_PPYear";
                _comm.Parameters.AddWithValue("_PP", _pp);
                _comm.Parameters.AddWithValue("_PPYear", _ppYear);
                SqlDataReader _reader = _comm.ExecuteReader();
                if (_reader.HasRows)
                {
                    _reader.Read();
                    _ret = _reader["StartDate"].ToString();
                }
                if (_reader.IsClosed != true) _reader.Close();
            }

            return _ret;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1)
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            string[] _sites = new string[] { "ACH", "FMC", "PLC", "RGH", "SPT", "SHC" };            

            try
            {
                using (var package = new ExcelPackage(new System.IO.FileInfo(Application.StartupPath + @"\Template.dat")))
                {
                    ExcelWorksheet ws;

                    string _PayPeriod = "Pay Period: " + cboPP.SelectedItem + "/" + cboYearPP.SelectedItem + "   Changes Effective " + GetStartPP(cboPP.SelectedItem.ToString(), cboYearPP.SelectedItem.ToString());

                    using (SqlConnection myConnection = new SqlConnection())
                    {
                        myConnection.ConnectionString = Common.SystemsServer;
                        myConnection.Open();

                        SqlCommand myCommand = myConnection.CreateCommand();

                        int _lineCtr;
                        int _row;

                        #region Export New Primary Positions to Excel

                        _lineCtr = 0;

                        ws = package.Workbook.Worksheets[1]; // New Primary Positions Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_NewPrimaryPositions " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 12 + _lineCtr; // 12 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["Unit"].ToString();
                                    ws.Cells[_row, 5].Value = _dr["Occupation"].ToString();
                                    ws.Cells[_row, 6].Value = _dr["Status"].ToString();
                                    _lineCtr++;
                                    //ws.InsertRow(27, 1, 28);
                                    //ws.Cells[27, 1].LoadFromText("45,76,12,1,darwin radwin 2 " + DateTime.Now.ToString("hh:mm:ss"));
                                }
                                _row = 12 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 12 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 12 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }
                        #endregion

                        #region Export Unit to Unit Transfer to Excel

                        _lineCtr = 0;                        

                        ws = package.Workbook.Worksheets[2]; // Unit to Unit Transfer Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_UnitToUnitTransfer " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 26 + _lineCtr; // 26 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["UnitFrom"].ToString();
                                    ws.Cells[_row, 5].Value = _dr["UnitTo"].ToString();
                                    ws.Cells[_row, 6].Value = _dr["Occupation"].ToString();
                                    ws.Cells[_row, 7].Value = _dr["ChangeInOccupation"].ToString().ToUpper() == "TRUE" ? "∆" : "";
                                    ws.Cells[_row, 8].Value = _dr["Status"].ToString();
                                    ws.Cells[_row, 10].Value = _dr["Comments"].ToString();

                                    if (_dr["ChangeInSite"].ToString().ToUpper() == "TRUE")
                                    {
                                        var range = ws.Cells[_row, 1, _row, 10];
                                        range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        range.Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                                    }

                                    _lineCtr++;
                                    //ws.InsertRow(27, 1, 28);
                                    //ws.Cells[27, 1].LoadFromText("45,76,12,1,darwin radwin 2 " + DateTime.Now.ToString("hh:mm:ss"));
                                }
                                _row = 26 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 26 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 26 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }

                        #endregion

                        #region Export Status Changes

                        _lineCtr = 0;

                        ws = package.Workbook.Worksheets[3]; // Status Changes Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_StatusChange " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 19 + _lineCtr; // 19 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["StatusFrom"].ToString();
                                    ws.Cells[_row, 5].Value = _dr["StatusTo"].ToString();
                                    ws.Cells[_row, 6].Value = _dr["Unit"].ToString();
                                    ws.Cells[_row, 7].Value = _dr["Comments"].ToString();
                                    _lineCtr++;                                   
                                }
                                _row = 19 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 19 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 19 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }
                        #endregion

                        #region Export Occupation Changes

                        _lineCtr = 0;

                        ws = package.Workbook.Worksheets[4]; // Occupation Changes Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_OccupationChange " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 23 + _lineCtr; // 23 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["Unit"].ToString();
                                    ws.Cells[_row, 5].Value = _dr["OccFrom"].ToString();
                                    ws.Cells[_row, 6].Value = _dr["OccTo"].ToString();
                                    ws.Cells[_row, 7].Value = _dr["Comments"].ToString();
                                    _lineCtr++;
                                }
                                _row = 23 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 23 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 23 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }
                        #endregion

                        #region Export Terminations

                        _lineCtr = 0;

                        ws = package.Workbook.Worksheets[5]; // Terminations Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_Terminations " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 19 + _lineCtr; // 19 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["Unit"].ToString();
                                    ws.Cells[_row, 5].Value = Convert.ToDateTime(_dr["TerminationDate"]).ToString("dd-MMM-yyyy");
                                    ws.Cells[_row, 6].Value = _dr["Comments"].ToString();
                                    _lineCtr++;
                                }
                                _row = 19 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 19 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 19 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }
                        #endregion

                        #region Export Transfers

                        _lineCtr = 0;

                        ws = package.Workbook.Worksheets[6]; // Transfers Sheet
                        ws.Cells[1, 1].Value = _PayPeriod; // set the payperiod and date on the sheet

                        for (int i = 1; i <= _sites.Length; i++)
                        {
                            myCommand.Parameters.Clear();

                            myCommand.CommandText = "select * From ItemsRpt_Transfers " +
                                "where PayPeriod = @_PayPeriod and PayPeriod_Year = @_PayPeriod_Year and ItemsReportLetter = @_IRL and [Site] = @_Site " +
                                "order by Emp_Name";

                            myCommand.Parameters.AddWithValue("_PayPeriod", cboPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", cboYearPP.SelectedItem);
                            myCommand.Parameters.AddWithValue("_IRL", cboItemsReport.SelectedItem);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            SqlDataReader _dr;
                            _dr = myCommand.ExecuteReader();

                            if (_dr.HasRows)
                            {
                                while (_dr.Read())
                                {
                                    _row = 26 + _lineCtr; // 26 is the starting line of the insert
                                    ws.InsertRow(_row, 1, _row + 1);
                                    ws.Cells[_row, 1].Value = _sites[i - 1];
                                    ws.Cells[_row, 2].Value = _dr["Emp_Num"].ToString();
                                    ws.Cells[_row, 3].Value = _dr["Emp_Name"].ToString();
                                    ws.Cells[_row, 4].Value = _dr["UnitFrom"].ToString();
                                    ws.Cells[_row, 5].Value = _dr["UnitTo"].ToString();
                                    ws.Cells[_row, 6].Value = _dr["Comments"].ToString();
                                    _lineCtr++;
                                }
                                _row = 26 + _lineCtr; ws.InsertRow(_row, 1, _row + 1); ws.Cells[_row, 1].Value = ""; _lineCtr++; // Insert a blank line after each site
                            }
                            else // no entries for the site
                            {
                                _row = 26 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = _sites[i - 1]; // Put the site name
                                _lineCtr++;
                                _row = 26 + _lineCtr;
                                ws.InsertRow(_row, 1, _row + 1);
                                ws.Cells[_row, 1].Value = "";
                                _lineCtr++; // Insert a blank line 
                            }
                            _dr.Close();
                        }
                        #endregion

                    }

                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 1;
                    saveFileDialog1.FileName = cboItemsReport.SelectedItem + ".xlsx";
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

        private void cboSite_NPP_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                Control p;
                p = ((Control)sender).Parent;
                p.SelectNextControl(ActiveControl, true, true, true, true);
            }
        }

        private void btnSave_UUT_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_UUT.SelectedIndex == -1 || txtEmpName_UUT.Text.Trim() == "" || txtTransFrom_UUT.Text.Trim() == "" || txtTransTo_UUT.Text.Trim() == "" || txtOcc_UUT.Text.Trim() == "" || txtStatus_UUT.Text.Trim() == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_UUT.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_UUT.Focus();
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (btnSave_UUT.Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_UnitToUnitTransfer (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, UnitFrom, UnitTo, Occupation, Status, ChangeInOccupation, ChangeInSite, Comments, EnteredBy) values " +
                            "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_UnitFrom, @_UnitTo, @_Occupation, @_Status, @_ChangeInOccupation, @_ChangeInSite, @_Comments, @_EnteredBy)";
                    }
                    else // btnSave_UUT.Text == "Update"
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "Update ItemsRpt_UnitToUnitTransfer SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, PayPeriod_Year = @_PayPeriod_Year, " +
                            "Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, UnitFrom = @_UnitFrom, UnitTo = @_UnitTo, Occupation = @_Occupation, Status = @_Status, " +
                            "ChangeInOccupation = @_ChangeInOccupation, ChangeInSite = @_ChangeInSite, Comments = @_Comments, EnteredBy = @_EnteredBy WHERE ID = " + ID;                            
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_UUT.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_UnitFrom", txtTransFrom_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_UnitTo", txtTransTo_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Occupation", txtOcc_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Status", txtStatus_UUT.Text.Trim().ToUpper());
                    myCommand.Parameters.AddWithValue("_ChangeInOccupation", chkChangeInOcc_UUT.Checked.ToString());
                    myCommand.Parameters.AddWithValue("_ChangeInSite", chkChangeInSite_UUT.Checked.ToString());
                    myCommand.Parameters.AddWithValue("_Comments", txtComments_UUT.Text.Trim());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 1, "UUT");                        
                    }

                    _frmReport.Show();
                    if (_frmReport.firstLoad) _frmReport.tabControl1.TabPages[1].Show();
                    _frmReport.Load_UUT_Data(_pp, _ppYear, _ItemsReportLetter);                    
                    _frmReport.tabControl1.SelectedIndex = 1;                   
                    ClearForm(tabControl1.TabPages[1]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void HideCancelBtn(Control _senderBtn, byte _tabNumber, string _tabName)
        {
            ((Button)_senderBtn).Text = "Save";
            var _CancelBtn = tabControl1.TabPages[_tabNumber].Controls.Find(((Button)_senderBtn).Tag.ToString(), true).FirstOrDefault();
            ((Button)_CancelBtn).Visible = false;
            ToggleTabs(true, _tabName);            
        }

        private void ClearForm(Control _parent)
        {
            foreach (Control c in _parent.Controls)
            {
                if (c.GetType() == typeof(TextBox))
                {
                    c.Text = "";
                }
                else if (c.GetType() == typeof(ComboBox))
                {
                    ((ComboBox)c).SelectedIndex = -1;                    
                }
                else if (c.GetType() == typeof(CheckBox))
                {
                    ((CheckBox)c).Checked = false;
                }
            }
        }

        private void ToggleTabs(bool _status, string _tabName)
        {            
            foreach (TabPage t in tabControl1.TabPages)
            {
                if (t.Name != _tabName) t.Enabled = _status;
                else { tabControl1.SelectedTab = t; }
            }
            cboPP.Enabled = cboYearPP.Enabled = cboItemsReport.Enabled = btnViewRpt.Enabled = _status;
        }

        private void btnSave_SC_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_SC.SelectedIndex == -1 || txtEmpNo_SC.Text.Trim() == "" || txtEmpName_SC.Text.Trim() == "" || txtStatusFrom.Text.Trim() == "" || txtStatusTo.Text.Trim() == "" || txtUnit_SC.Text.Trim() == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_SC.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_SC.Focus();
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (((Button)sender).Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_StatusChange (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, StatusFrom, StatusTo, Unit, Comments, EnteredBy) values " +
                        "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_StatusFrom, @_StatusTo, @_Unit, @_Comments, @_EnteredBy)";
                    }
                    else // Update
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "UPDATE ItemsRpt_StatusChange SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, PayPeriod_Year = @_PayPeriod_Year, " +
                            "Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, StatusFrom = @_StatusFrom, StatusTo = @_StatusTo, Unit = @_Unit, Comments = @_Comments,  " +
                            "EnteredBy = @_EnteredBy WHERE ID = " + ID;                        
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_SC.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_SC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_SC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_StatusFrom", txtStatusFrom.Text.Trim().ToUpper());
                    myCommand.Parameters.AddWithValue("_StatusTo", txtStatusTo.Text.Trim().ToUpper());
                    myCommand.Parameters.AddWithValue("_Unit", txtUnit_SC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Comments", txtComment_SC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);                    

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 2, "SC");                        
                    }

                    _frmReport.Load_SC_Data(_pp, _ppYear, _ItemsReportLetter);
                    _frmReport.Show();
                    _frmReport.tabControl1.SelectedIndex = 2;
                    ClearForm(tabControl1.TabPages[2]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void btnSave_OC_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_OC.SelectedIndex == -1 || txtEmpNo_OC.Text.Trim() == "" || txtEmpName_OC.Text.Trim() == "" || txtUnit_OC.Text.Trim() == "" || txtOccFrom_OC.Text.Trim() == "" || txtOccTo_OC.Text.Trim() == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_OC.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_OC.Focus();
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (((Button)sender).Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_OccupationChange (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, Unit, OccFrom, OccTo, Comments, EnteredBy) values " +
                        "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_Unit, @_OccFrom, @_OccTo, @_Comments, @_EnteredBy)";
                    }
                    else // Update
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "UPDATE ItemsRpt_OccupationChange SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, PayPeriod_Year = @_PayPeriod_Year, " +
                            "Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, Unit = @_Unit, OccFrom = @_OccFrom, OccTo = @_OccTo, Comments = @_Comments, " +
                            "EnteredBy = @_EnteredBy WHERE ID = " + ID;
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_OC.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Unit", txtUnit_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_OccFrom", txtOccFrom_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_OccTo", txtOccTo_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Comments", txtComments_OC.Text.Trim());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 3, "OC");                        
                    }

                    _frmReport.Load_OC_Data(_pp, _ppYear, _ItemsReportLetter);
                    _frmReport.Show();
                    _frmReport.tabControl1.SelectedIndex = 3;
                    ClearForm(tabControl1.TabPages[3]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void btnSave_Terms_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_Terms.SelectedIndex == -1 || txtEmpNo_Terms.Text.Trim() == "" || txtEmpName_Terms.Text.Trim() == "" || txtUnit_Terms.Text.Trim() == "" || txtComments_Terms.Text.Trim() == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_Terms.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_Terms.Focus();
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (((Button)sender).Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_Terminations (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, Unit, TerminationDate, Comments, EnteredBy) values " +
                        "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_Unit, @_TerminationDate, @_Comments, @_EnteredBy)";
                    }
                    else // Update
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "UPDATE ItemsRpt_Terminations SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, PayPeriod_Year = @_PayPeriod_Year, " +
                            "Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, Unit = @_Unit, TerminationDate = @_TerminationDate, Comments = @_Comments, EnteredBy = @_EnteredBy " +
                            "WHERE ID = " + ID;
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_Terms.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_Terms.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_Terms.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Unit", txtUnit_Terms.Text.Trim());
                    myCommand.Parameters.AddWithValue("_TerminationDate", dp_Terms.Value.ToString("yyyy-MM-dd"));
                    myCommand.Parameters.AddWithValue("_Comments", txtComments_Terms.Text.Trim());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 4, "Terms");                        
                    }

                    _frmReport.Load_Terms_Data(_pp, _ppYear, _ItemsReportLetter);
                    _frmReport.Show();
                    _frmReport.tabControl1.SelectedIndex = 4;
                    ClearForm(tabControl1.TabPages[4]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void btnSave_Trans_Click(object sender, EventArgs e)
        {
            if (((Button)sender).Text == "Save" && (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1))
            {
                MessageBox.Show("Please select payperiod or year or the items report.");
                return;
            }

            if (cboSite_Trans.SelectedIndex == -1 || txtEmpNo_Trans.Text.Trim() == "" || txtEmpName_Trans.Text.Trim() == "" || txtUnitFrom_Trans.Text.Trim() == "" || txtUnitTo_Trans.Text.Trim() == "" || txtComments_Trans.Text == "")
            {
                MessageBox.Show("Please check again your inputs, blank field detected.");
                return;
            }

            if (txtEmpNo_Trans.Text.Trim().Length != 10)
            {
                MessageBox.Show("Please check again employee number, it should include the record number.");
                txtEmpNo_Trans.Focus();
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    string _pp;
                    string _ppYear;
                    string _ItemsReportLetter;

                    if (((Button)sender).Text == "Save")
                    {
                        _pp = cboPP.SelectedItem.ToString();
                        _ppYear = cboYearPP.SelectedItem.ToString();
                        _ItemsReportLetter = cboItemsReport.SelectedItem.ToString();

                        myCommand.CommandText = "Insert into ItemsRpt_Transfers (ItemsReportLetter, PayPeriod, PayPeriod_Year, Site, Emp_Num, Emp_Name, UnitFrom, UnitTo, Comments, EnteredBy) values " +
                        "(@_ItemsReportLetter, @_PayPeriod, @_PayPeriod_Year, @_Site, @_Emp_Num, @_Emp_Name, @_UnitFrom, @_UnitTo, @_Comments, @_EnteredBy)";
                    }
                    else // Update
                    {
                        _pp = pp;
                        _ppYear = ppYear;
                        _ItemsReportLetter = itemsReportLetter;

                        myCommand.CommandText = "UPDATE ItemsRpt_Transfers SET ItemsReportLetter = @_ItemsReportLetter, PayPeriod = @_PayPeriod, PayPeriod_Year = @_PayPeriod_Year, " +
                            "Site = @_Site, Emp_Num = @_Emp_Num, Emp_Name = @_Emp_Name, UnitFrom = @_UnitFrom, UnitTo = @_UnitTo, Comments = @_Comments, " +
                            "EnteredBy = @_EnteredBy WHERE ID = " + ID;
                    }

                    myCommand.Parameters.AddWithValue("_ItemsReportLetter", _ItemsReportLetter);
                    myCommand.Parameters.AddWithValue("_PayPeriod", _pp);
                    myCommand.Parameters.AddWithValue("_PayPeriod_Year", _ppYear);
                    myCommand.Parameters.AddWithValue("_Site", (cboSite_Trans.SelectedIndex + 1).ToString());
                    myCommand.Parameters.AddWithValue("_Emp_Num", txtEmpNo_Trans.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Emp_Name", txtEmpName_Trans.Text.Trim());
                    myCommand.Parameters.AddWithValue("_UnitFrom", txtUnitFrom_Trans.Text.Trim());
                    myCommand.Parameters.AddWithValue("_UnitTo", txtUnitTo_Trans.Text.Trim());
                    myCommand.Parameters.AddWithValue("_Comments", txtComments_Trans.Text.Trim());
                    myCommand.Parameters.AddWithValue("_EnteredBy", Common.CurrentUser);

                    myCommand.ExecuteNonQuery();
                    myCommand.Dispose();

                    if (((Button)sender).Text == "Save")
                    {
                        MessageBox.Show("Successfully Saved!", "Confirmation");
                    }
                    else
                    {
                        MessageBox.Show("Successfully Updated!", "Confirmation");
                        HideCancelBtn((Control)sender, 5, "Trans");
                        _frmReport.Load_Trans_Data(_pp, _ppYear, _ItemsReportLetter);
                        _frmReport.Show();
                    }
                    
                    _frmReport.Load_Trans_Data(_pp, _ppYear, _ItemsReportLetter);
                    _frmReport.Show();
                    _frmReport.tabControl1.SelectedIndex = 5;
                    ClearForm(tabControl1.TabPages[5]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (cboPP.SelectedIndex == -1 || cboYearPP.SelectedIndex == -1 || cboItemsReport.SelectedIndex == -1)
                {
                    MessageBox.Show("Please select payperiod or year or the items report.");
                    return;
                }

                _frmReport.cboPP.SelectedItem = cboPP.SelectedItem;
                _frmReport.cboYearPP.SelectedItem = cboYearPP.SelectedItem;
                _frmReport.cboItemsReport.SelectedItem = cboItemsReport.SelectedItem;
                _frmReport.Show();
                _frmReport.LoadAllData();
                _frmReport.Focus();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "; Error in loading defaults.");
            }
        }

        private void btnCancel_UUT_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_UUT, 1, "UUT");            
            ClearForm(tabControl1.TabPages[1]);
            _frmReport.Show();
        }

        public void Load_UUT_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; //@"Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=" + Application.StartupPath + @"\items.mdb;Uid=Admin;Pwd=;";
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.ID, U.Emp_Num, U.Emp_Name, U.UnitFrom, U.UnitTo, U.Occupation, U.ChangeInOccupation, " +
                        "U.Status, U.Comments, U.EnteredBy, U.ChangeInSite from ItemsRpt_UnitToUnitTransfer U where ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);                    

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_UUT.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_UUT.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_UUT.Text = _dr["Emp_Name"].ToString();
                        txtTransFrom_UUT.Text = _dr["UnitFrom"].ToString();
                        txtTransTo_UUT.Text = _dr["UnitTo"].ToString();
                        txtOcc_UUT.Text = _dr["Occupation"].ToString();
                        txtStatus_UUT.Text = _dr["Status"].ToString();
                        chkChangeInOcc_UUT.Checked = Convert.ToBoolean(_dr["ChangeInOccupation"].ToString());
                        chkChangeInSite_UUT.Checked = Convert.ToBoolean(_dr["ChangeInSite"].ToString());
                        txtComments_UUT.Text = _dr["Comments"].ToString();

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_UUT.Text = "Update";
                        btnCancel_UUT.Visible = true;
                        ToggleTabs(false, "UUT");
                    }
                    myCommand.Dispose();                    
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        public void Load_NPP_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer; 
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.ID, U.Emp_Num, U.Emp_Name, U.Unit, U.Occupation, U.Status " +
                        "from ItemsRpt_NewPrimaryPositions U where U.ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_NPP.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_NPP.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_NPP.Text = _dr["Emp_Name"].ToString();
                        txtUnit_NPP.Text = _dr["Unit"].ToString();
                        txtOcc_NPP.Text = _dr["Occupation"].ToString();
                        txtStatus_NPP.Text = _dr["Status"].ToString();                        

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_NPP.Text = "Update";
                        btnCancel_NPP.Visible = true;
                        ToggleTabs(false, "NPP");
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        public void Load_SC_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.Emp_Num, U.Emp_Name, U.StatusFrom, U.StatusTo, U.Unit, " +
                        "U.Comments from ItemsRpt_StatusChange U where U.ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_SC.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_SC.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_SC.Text = _dr["Emp_Name"].ToString();
                        txtStatusFrom.Text = _dr["StatusFrom"].ToString();
                        txtStatusTo.Text = _dr["StatusTo"].ToString();
                        txtUnit_SC.Text = _dr["Unit"].ToString();
                        txtComment_SC.Text = _dr["Comments"].ToString();

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_SC.Text = "Update";
                        btnCancel_SC.Visible = true;
                        ToggleTabs(false, "SC");
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        public void Load_OC_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.Emp_Num, U.Emp_Name, U.Unit, U.OccFrom, U.OccTo, " +
                        "U.Comments from ItemsRpt_OccupationChange U where U.ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_OC.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_OC.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_OC.Text = _dr["Emp_Name"].ToString();
                        txtUnit_OC.Text = _dr["Unit"].ToString();
                        txtOccFrom_OC.Text = _dr["OccFrom"].ToString();
                        txtOccTo_OC.Text = _dr["OccTo"].ToString();
                        txtComments_OC.Text = _dr["Comments"].ToString();

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_OC.Text = "Update";
                        btnCancel_OC.Visible = true;
                        ToggleTabs(false, "OC");
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        public void Load_Terms_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.Emp_Num, U.Emp_Name, U.Unit, U.TerminationDate, " +
                        "U.Comments from ItemsRpt_Terminations U where U.ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_Terms.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_Terms.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_Terms.Text = _dr["Emp_Name"].ToString();
                        txtUnit_Terms.Text = _dr["Unit"].ToString();
                        dp_Terms.Value = Convert.ToDateTime(_dr["TerminationDate"]);                        
                        txtComments_Terms.Text = _dr["Comments"].ToString();

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_Terms.Text = "Update";
                        btnCancel_Terms.Visible = true;
                        ToggleTabs(false, "Terms");
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        public void Load_Trans_Data(string _ID)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "Select U.Site, U.PayPeriod, U.PayPeriod_Year, U.ItemsReportLetter, U.Emp_Num, U.Emp_Name, U.UnitFrom, U.UnitTo, " +
                        "U.Comments from ItemsRpt_Transfers U where U.ID = @_ID";
                    myCommand.Parameters.AddWithValue("_ID", _ID);

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboSite_Trans.SelectedIndex = Convert.ToInt32(_dr["Site"]) - 1;
                        txtEmpNo_Trans.Text = _dr["Emp_Num"].ToString();
                        txtEmpName_Trans.Text = _dr["Emp_Name"].ToString();
                        txtUnitFrom_Trans.Text = _dr["UnitFrom"].ToString();
                        txtUnitTo_Trans.Text = _dr["UnitTo"].ToString();
                        txtComments_Trans.Text = _dr["Comments"].ToString();

                        pp = _dr["PayPeriod"].ToString();
                        ppYear = _dr["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr["ItemsReportLetter"].ToString();

                        btnSave_Trans.Text = "Update";
                        btnCancel_Trans.Visible = true;
                        ToggleTabs(false, "Trans");
                    }
                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void ItemsReport_FormClosing(object sender, FormClosingEventArgs e)
        {
            _frmReport.Close();            
        }

        private void btnCancel_NPP_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_NPP, 0, "NPP");
            ClearForm(tabControl1.TabPages[0]);
            _frmReport.Show();
        }

        private void btnCancel_SC_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_SC, 2, "SC");
            ClearForm(tabControl1.TabPages[2]);
            _frmReport.Show();
        }

        private void btnCancel_OC_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_OC, 3, "OC");
            ClearForm(tabControl1.TabPages[3]);
            _frmReport.Show();
        }

        private void btnCancel_Terms_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_Terms, 4, "Terms");
            ClearForm(tabControl1.TabPages[4]);
            _frmReport.Show();
        }

        private void btnCancel_Trans_Click(object sender, EventArgs e)
        {
            HideCancelBtn(btnSave_Trans, 5, "Trans");
            ClearForm(tabControl1.TabPages[5]);
            _frmReport.Show();
        }

        private void txtUnit_NPP_Leave(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Trim() == "") return;
            var _cboSite = tabControl1.TabPages[tabControl1.SelectedIndex].Controls.Find(((TextBox)sender).Tag.ToString(), true).FirstOrDefault();
            ((ComboBox)_cboSite).SelectedIndex = GetSiteNum_ShortDesc(((TextBox)sender).Text.Trim().ToUpper());            
        }

        private void txtUnit_Terms_Leave(object sender, EventArgs e)
        {
            if (((TextBox)sender).Text.Trim() == "") return;
            var _cboSite = tabControl1.TabPages[tabControl1.SelectedIndex].Controls.Find(((TextBox)sender).Tag.ToString(), true).FirstOrDefault();
            ((ComboBox)_cboSite).SelectedIndex = GetSiteNum_LongDesc(((TextBox)sender).Text.Trim().ToUpper());
        }

        private void txtEmpNo_UUT_KeyPress(object sender, KeyPressEventArgs e)
        {

        }
    }
}
