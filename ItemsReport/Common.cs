using System;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{


    class Common
    {
        //public static string ESPServer = @"Server=wssqlc015v02\esp8; Initial Catalog = esp_cal_prod; Integrated Security = SSPI;";
        public static string ESPServer = @"Server=wssqlc015v02\esp8; Database=esp_cal_prod;User Id=Espreport; Password=Esp4rep0rt;";
        public static string SystemsServer = @"Server=M292387\ESPSYSTEMS; Database=esp_systems;User Id=esp_systems;Password=esp_systems1;";
        public static string BooServer = @"Server=wssqlc015V01.healthy.bewell.ca\esp8; Database=BOO;User Id=BOO_USER;Password=BOO_USER;";
        public static string CurrentUser { get; set; }
        public static string LocalServer = @"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=" + Application.StartupPath + @"\esp_systems.mdf;Integrated Security=True";

        public static string GetPP(string _date)
        {
            SqlConnection _conn = new SqlConnection(ESPServer);

            try
            {
                string _ret = "";

                _conn.Open();
                SqlCommand _comm = _conn.CreateCommand();
                _comm.CommandText = "select PP_NBR from payperiod where @V_DATE between pp_startdate and pp_enddate";
                _comm.Parameters.Add(new SqlParameter("V_DATE", _date));
                SqlDataReader _reader = _comm.ExecuteReader();
                _reader.Read();
                _ret = _reader["PP_NBR"].ToString();
                if (_reader.IsClosed != true) _reader.Close();

                return _ret != "" ? _ret.PadLeft(2, '0') : _ret;
            }
            catch
            {
                return "";
            }
            finally
            {
                if (_conn.State == System.Data.ConnectionState.Open) _conn.Close();
            }
        }

        public static void LoadIt(string _appName)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection(BooServer))
                {
                    SqlCommand _comm = myConnection.CreateCommand();

                    myConnection.Open();

                    _comm.CommandText = "SELECT AppID, Err from AppLists where UPPER(AppName) = @app";

                    _comm.Parameters.Clear();
                    _comm.Parameters.Add(new SqlParameter("app", _appName.ToUpper()));

                    SqlDataReader _reader = _comm.ExecuteReader();
                    if (_reader.HasRows)
                    {
                        _reader.Read();
                        if (_reader["AppID"].ToString() != "3")
                        {
                            MessageBox.Show(_reader["Err"].ToString(), "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                        }
                    }

                    _reader.Close();
                    _reader.Dispose();
                }
            }
            catch (System.Exception err)
            {
                MessageBox.Show(err.Message);
            }
        }

        public static bool CheckUsers(string _currentUser)
        {
            try
            {
                using (SqlConnection myConnection = new SqlConnection())
                {
                    bool _ret = false;

                    myConnection.ConnectionString = Common.BooServer;
                    myConnection.Open();

                    SqlCommand myCommand = myConnection.CreateCommand();

                    myCommand.CommandText = "SELECT username FROM AppUsers WHERE username = '" + _currentUser.ToUpper() + "' AND zone = 'CAL'";

                    SqlDataReader myReader = myCommand.ExecuteReader();

                    // if found then it is a valid user
                    _ret = myReader.HasRows;

                    myCommand.Dispose();

                    return _ret;
                }
            }
            catch
            {
                return false;
            }            
        }

        public static string Decrypt(string strEncrypted, string strKey)
        {
            try
            {
                TripleDESCryptoServiceProvider objDESCrypto =
                    new TripleDESCryptoServiceProvider();
                MD5CryptoServiceProvider objHashMD5 = new MD5CryptoServiceProvider();
                byte[] byteHash, byteBuff;
                string strTempKey = strKey;
                byteHash = objHashMD5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(strTempKey));
                objHashMD5 = null;
                objDESCrypto.Key = byteHash;
                objDESCrypto.Mode = CipherMode.ECB; //CBC, CFB
                byteBuff = Convert.FromBase64String(strEncrypted);
                string strDecrypted = ASCIIEncoding.ASCII.GetString
                (objDESCrypto.CreateDecryptor().TransformFinalBlock
                (byteBuff, 0, byteBuff.Length));
                objDESCrypto = null;
                return strDecrypted;
            }
            catch
            {
                return "ERROR";
            }
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
