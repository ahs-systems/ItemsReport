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
    public partial class frmWorkingStatus: Form
    {
        private Button btnRefresh;
        private ComboBox cboWorkingStatus;
        private RichTextBox rtbWorkers;
        private Label label1;
        private bool firstLoad = true;

        public frmWorkingStatus()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmWorkingStatus));
            this.rtbWorkers = new System.Windows.Forms.RichTextBox();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.cboWorkingStatus = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // rtbWorkers
            // 
            this.rtbWorkers.BackColor = System.Drawing.Color.WhiteSmoke;
            this.rtbWorkers.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtbWorkers.Location = new System.Drawing.Point(12, 90);
            this.rtbWorkers.Name = "rtbWorkers";
            this.rtbWorkers.ReadOnly = true;
            this.rtbWorkers.Size = new System.Drawing.Size(391, 145);
            this.rtbWorkers.TabIndex = 3;
            this.rtbWorkers.Text = "";
            // 
            // btnRefresh
            // 
            this.btnRefresh.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.btnRefresh.Image = ((System.Drawing.Image)(resources.GetObject("btnRefresh.Image")));
            this.btnRefresh.Location = new System.Drawing.Point(12, 39);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(104, 40);
            this.btnRefresh.TabIndex = 2;
            this.btnRefresh.Text = "Refresh";
            this.btnRefresh.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText;
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // cboWorkingStatus
            // 
            this.cboWorkingStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboWorkingStatus.Font = new System.Drawing.Font("Verdana", 7.471698F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboWorkingStatus.FormattingEnabled = true;
            this.cboWorkingStatus.Items.AddRange(new object[] {
            "I\'m Not Working On It (Yehey!)",
            "I\'m Working On It",
            "I\'m Done Working On It"});
            this.cboWorkingStatus.Location = new System.Drawing.Point(12, 12);
            this.cboWorkingStatus.Name = "cboWorkingStatus";
            this.cboWorkingStatus.Size = new System.Drawing.Size(242, 21);
            this.cboWorkingStatus.TabIndex = 1;
            this.cboWorkingStatus.SelectedIndexChanged += new System.EventHandler(this.cboWorkingStatus_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Verdana", 8.150944F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.ForeColor = System.Drawing.Color.Maroon;
            this.label1.Location = new System.Drawing.Point(12, 238);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(391, 34);
            this.label1.TabIndex = 4;
            this.label1.Text = "Note: The last person working on it is the one who usually sends the \'Items Repor" +
    "t\' to SSO.";
            // 
            // frmWorkingStatus
            // 
            this.ClientSize = new System.Drawing.Size(415, 274);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboWorkingStatus);
            this.Controls.Add(this.btnRefresh);
            this.Controls.Add(this.rtbWorkers);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmWorkingStatus";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Work Status";
            this.Load += new System.EventHandler(this.frmWorkingStatus_Load);
            this.ResumeLayout(false);

        }

        private void frmWorkingStatus_Load(object sender, EventArgs e)
        {
            CheckStatus();
        }

        private void CheckStatus()
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    // Get your current working status
                    myCommand.CommandText = "select * from ItemsRpt_WorkStatus where wName = @_name and workingDate = @_date";
                    myCommand.Parameters.AddWithValue("_name", Common.CurrentUser);
                    myCommand.Parameters.AddWithValue("_date", DateTime.Today.ToString("dd-MMM-yyyy"));

                    SqlDataReader _dr = myCommand.ExecuteReader();

                    if (_dr.HasRows)
                    {
                        _dr.Read();
                        cboWorkingStatus.SelectedIndex = Convert.ToInt16(_dr["wStatus"]);
                    }
                    else
                    {
                        firstLoad = false;
                    }

                    
                    _dr.Close();

                    // Get the current workers working on the Items Report
                    rtbWorkers.Clear();

                    myCommand.Parameters.Clear();
                    myCommand.CommandText = "select * from ItemsRpt_WorkStatus where workingDate = @_date order by dateUpdated";
                    myCommand.Parameters.AddWithValue("_date", DateTime.Today.ToString("dd-MMM-yyyy"));
                    _dr = myCommand.ExecuteReader();
                    if (_dr.HasRows)
                    {
                        //bool thereAreStillWorkingOnIt = false;
                        //string _lastWorker = "";
                        while (_dr.Read())
                        {
                            if (Convert.ToByte(_dr["wStatus"]) == (byte) WorkingStatus.WorkingOnIt)
                            {
                                rtbWorkers.AppendText("[As of " + Convert.ToDateTime(_dr["dateUpdated"]).ToString("hh:mm:ss tt") + "] ", Color.DeepSkyBlue);
                                rtbWorkers.AppendText(_dr["wName"].ToString().Replace(@"HEALTHY\",""), Color.Red, true);
                                rtbWorkers.AppendText(" is still working on it.", Color.DimGray, false);
                                rtbWorkers.AppendText(Environment.NewLine);

                                //thereAreStillWorkingOnIt = true;
                            }
                            else if (Convert.ToByte(_dr["wStatus"]) == (byte)WorkingStatus.DoneWorkingOnIt)
                            {
                                rtbWorkers.AppendText("[As of " + Convert.ToDateTime(_dr["dateUpdated"]).ToString("hh:mm:ss tt") + "] ", Color.DeepSkyBlue);
                                rtbWorkers.AppendText(_dr["wName"].ToString().Replace(@"HEALTHY\", ""), Color.Green, true);
                                rtbWorkers.AppendText(" is done working on it.", Color.DimGray, false);
                                rtbWorkers.AppendText(Environment.NewLine);
                            }
                            else 
                            {
                                rtbWorkers.AppendText("[As of " + Convert.ToDateTime(_dr["dateUpdated"]).ToString("hh:mm:ss tt") + "] ", Color.DarkGray);
                                rtbWorkers.AppendText(_dr["wName"].ToString().Replace(@"HEALTHY\", ""), Color.DarkGray, true);
                                rtbWorkers.AppendText(" is not working on it.", Color.DarkGray, false);
                                rtbWorkers.AppendText(Environment.NewLine);
                            }

                            //_lastWorker = _dr["wName"].ToString();
                        }
                        //if (_lastWorker == Common.CurrentUser && !thereAreStillWorkingOnIt)
                        //{
                        //    MessageBox.Show("It seems that you're the last person working on it.\n\nUsually the last person is the one who will send the 'Items Report' to SSO. \n\nThank you!","Message",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                        //}
                    }                    
                    _dr.Close();


                    myCommand.Dispose();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void cboWorkingStatus_SelectedIndexChanged(object sender, EventArgs e)
        {
            // If form is first time loaded, dont update the timestamp of the current working status
            if (firstLoad)
            {
                firstLoad = false;
                return;
            }

            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    myConnection.ConnectionString = Common.SystemsServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "if exists (SELECT * FROM ItemsRpt_WorkStatus WHERE workingDate = @_workingDate and wName = @_wName) " +
                                    "begin " +
                                    "    UPDATE ItemsRpt_WorkStatus SET wStatus = @_wStatus, dateUpdated = sysdatetime() WHERE workingDate = @_workingDate and wName = @_wName " +
                                    "end " +
                                    "else " +
                                    "begin " +
                                    "    INSERT INTO ItemsRpt_WorkStatus(wName, wStatus, workingDate) VALUES(@_wName, @_wStatus, @_workingDate) " +
                                    "end";
                    myCommand.Parameters.AddWithValue("_workingDate", DateTime.Today.ToString("dd-MMM-yyyy"));
                    myCommand.Parameters.AddWithValue("_wName", Common.CurrentUser);
                    myCommand.Parameters.AddWithValue("_wStatus", cboWorkingStatus.SelectedIndex);

                    myCommand.ExecuteNonQuery();                   
                    myCommand.Dispose();

                    CheckStatus();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ooops, there's an error: " + ex.Message, "ERROR");
            }
        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            CheckStatus();
        }
    }

    public static class RichTextBoxExtensions
    {
        public static void AppendText(this RichTextBox box, string text, Color color, bool _bold)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;
            box.SelectionColor = color;
            if (_bold)
            {
                box.SelectionFont = new Font(box.Font, FontStyle.Bold);
            }
            else
            {
                box.SelectionFont = new Font(box.Font, FontStyle.Regular);
            }
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }

        public static void AppendText(this RichTextBox box, string text, Color color)
        {
            box.SelectionStart = box.TextLength;
            box.SelectionLength = 0;
            box.SelectionColor = color;            
            box.AppendText(text);
            box.SelectionColor = box.ForeColor;
        }
    }
}
