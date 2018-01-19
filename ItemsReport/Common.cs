using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{


    class Common
    {
        //public static string ESPServer = @"Server=wssqlc015v02\esp8; Initial Catalog = esp_cal_prod; Integrated Security = SSPI;";
        public static string ESPServer = @"Server=wssqlc015v02\esp8; Database=esp_cal_prod;User Id=Espreport; Password=Esp4rep0rt;";
        public static string SystemsServer = @"Server=M292387\ESPSYSTEMS; Database=esp_systems;User Id=esp_systems;Password=esp_systems1;";
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
