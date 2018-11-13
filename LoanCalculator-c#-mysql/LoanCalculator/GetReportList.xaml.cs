using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using MySql.Data.MySqlClient;
using System.Data;
//using System.Windows.Forms;
namespace LoanCalculator
{
    /// <summary>
    /// GetReportList.xaml 的交互逻辑
    /// </summary>

    public class PassValuesEventArgs : EventArgs
    {
        public int Value { get; internal set; }
        public PassValuesEventArgs(int Value)
        {
            this.Value = Value;
        }
    }

    public class Emp
    {
        public String Name { get; set; }
        public int Id { get; set; }
    }
    public partial class GetReportList : Window
    {
        //EventArgs PassValuesEventArgs = new EventArgs();
        public System.Windows.Forms.ListViewItem.ListViewSubItemCollection SubItems { get; }
        public delegate void PassValuesHandler(object sender, PassValuesEventArgs e);
        public event PassValuesHandler PassValuesEvent;
        public GetReportList()
        {
            InitializeComponent();
            GetList();
        }
        public void GetList()
        {
            string connstring = "server=localhost;database=loancalculator;uid=root;pwd=root";
            MySqlConnection conn = new MySqlConnection(connstring);
            try
            { 
                conn.Open();
                //MessageBox.Show("连接成功", "测试结果");
            }
            catch (MySqlException ex)
            {
               // MessageBox.Show(ex.Message);
            }
            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = "select reportid,ReportName from annualreport";
            cmd.ExecuteNonQuery();
            MySqlDataAdapter adap = new MySqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            adap.Fill(ds);

            foreach (DataRow r in ds.Tables[0].Rows)
            {
                string Name = r["reportname"].ToString();
                string Id = r["reportid"].ToString();
                //MessageBox.Show(Name);
                lv_ReportList.Items.Add(new { Name = Name, Id = Id});
            }
            conn.Close();
        }
        /*
        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            string value1 = tbValue1.Text;   // Text Property return value is string type .
            int value2;
            Int32.TryParse(tbValue2.Text, out value2);

            PassValuesEventArgs args = new PassValuesEventArgs(value1, value2);
            PassValuesEvent(this, args);

            this.Close();
        }
        */
        public String getReportid(String a)
        {
            char ex = ' ';
            bool flag = false;
            String ans = "";
            for (int i = 0; i <= a.Length; ++i)
            {
                if (!flag) {
                    if (ex == 'I' && a[i] == 'd')
                    {
                        i+=3;
                        flag = true;
                    }
                    else ex = a[i];
                } else
                {
                    if (a[i] == ' ' || a[i]=='}') return ans;
                    ans += a[i];
                }
                    
            }
            return ans;
        }
        private void bt_Import_Click(object sender, RoutedEventArgs e)
        {
            //ListViewItem list = lv_ReportList；
            //String id = lv_ReportList.SelectedValue.ToString();
            int id = int.Parse(getReportid(lv_ReportList.SelectedValue.ToString()));
            PassValuesEventArgs args = new PassValuesEventArgs(id);
            PassValuesEvent(this, args);
            MainWindow f1 = (MainWindow)this.Owner;
            f1.RefreshMethod();
            this.Close();
        }

        private void bt_Delete_Click(object sender, RoutedEventArgs e)
        {
            int id = int.Parse(getReportid(lv_ReportList.SelectedValue.ToString()));
            lv_ReportList.Items.Clear();
            string connstring = "server=localhost;database=loancalculator;uid=root;pwd=root";
            MySqlConnection conn = new MySqlConnection(connstring);
            try
            {
                conn.Open();
                //MessageBox.Show("连接成功", "测试结果");
            }
            catch (MySqlException ex)
            {
                // MessageBox.Show(ex.Message);
            }
            MySqlCommand cmd = conn.CreateCommand();
            cmd.CommandText = "delete from peryearreport where sonid = " +
                "(select sonid1 from annualreport where reportid = @reportid)";
            cmd.Parameters.AddWithValue("@reportid", id);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            cmd = conn.CreateCommand();
            cmd.CommandText = "delete from peryearreport where sonid = " +
                "(select sonid2 from annualreport where reportid = @reportid)";
            cmd.Parameters.AddWithValue("@reportid", id);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            cmd = conn.CreateCommand();
            cmd.CommandText = "delete from peryearreport where sonid = " +
                "(select sonid3 from annualreport where reportid = @reportid)";
            cmd.Parameters.AddWithValue("@reportid", id);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            cmd = conn.CreateCommand();
            cmd.CommandText = "delete from annualreport where reportid = @reportid";
            cmd.Parameters.AddWithValue("@reportid", id);
            cmd.ExecuteNonQuery();
            cmd.Cancel();
            GetList();
        }
    }
}
