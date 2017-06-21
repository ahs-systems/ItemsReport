using System;
using System.Drawing;
using OfficeOpenXml;
using System.Data.SqlClient;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {

        public static string ESPServer = @"Server=wssqlc015v02\esp8; Initial Catalog = esp_cal_prod; Integrated Security = SSPI;";
        public static string SystemsServer = @"Server=M292387\ESPSYSTEMS; Database=esp_systems;User Id=esp_systems;Password=esp_systems1;";

        static void Main(string[] args)
        {
            Console.WriteLine("Generating Items Report Excel File...");
            CreateExcel();
            Console.WriteLine("Done.");
        }

        private static string GetStartPP(string _pp, string _ppYear)
        {
            string _ret = "";

            using (SqlConnection myConnection = new SqlConnection())
            {
                myConnection.ConnectionString = ESPServer;
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

        private static void CreateExcel()
        {
            string[] _sites = new string[] { "ACH", "FMC", "PLC", "RGH", "SPT", "SHC" };

            try
            {
                using (var package = new ExcelPackage(new System.IO.FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"\Template.dat")))
                {
                    ExcelWorksheet ws;                    

                    using (SqlConnection myConnection = new SqlConnection())
                    {
                        myConnection.ConnectionString = SystemsServer;
                        myConnection.Open();

                        SqlCommand myCommand = myConnection.CreateCommand();

                        string pp;
                        string ppYear;
                        string itemsReportLetter;

                        myCommand.CommandText = "select PayPeriod, PayPeriod_Year, ItemsReportLetter from ItemsRpt_UnitToUnitTransfer " +
                                                "where EnteredDate = (select max(EnteredDate) from ItemsRpt_UnitToUnitTransfer)";

                        SqlDataReader _dr1 = myCommand.ExecuteReader();
                        _dr1.Read();
                        pp = _dr1["PayPeriod"].ToString();
                        ppYear = _dr1["PayPeriod_Year"].ToString();
                        itemsReportLetter = _dr1["ItemsReportLetter"].ToString();

                        _dr1.Close();                        

                        string _PayPeriod = "Pay Period: " + pp + "/" + ppYear + "   Changes Effective " + GetStartPP(pp, ppYear);

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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
                            myCommand.Parameters.AddWithValue("_Site", i);

                            
                            SqlDataReader _dr = myCommand.ExecuteReader();

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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
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

                            myCommand.Parameters.AddWithValue("_PayPeriod", pp);
                            myCommand.Parameters.AddWithValue("_PayPeriod_Year", ppYear);
                            myCommand.Parameters.AddWithValue("_IRL", itemsReportLetter);
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

                    package.SaveAs(new FileInfo(@"\\jeeves.crha-health.ab.ca\rsss_systems\Operations - RSSS Systems Group\Automated Files\ItemsReport.xlsx"));
                    
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
}
