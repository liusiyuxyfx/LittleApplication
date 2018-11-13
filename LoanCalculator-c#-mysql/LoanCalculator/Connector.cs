using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Windows;
using System.Data.SqlClient;
using System.Data;
namespace LoanCalculator
{
    class Connector
    {
        MySqlConnection conn;
        public void connect()
        {
            string connstring = "server=localhost;database=loancalculator;uid=root;pwd=root";
            conn = new MySqlConnection(connstring);
            try
            {
                conn.Open();
                MessageBox.Show("连接成功", "测试结果");
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
            }
     
        }
    }
}
