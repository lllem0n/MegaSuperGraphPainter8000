using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MySql.Data.MySqlClient;
using System.Data;

namespace ZubApproxProject
{
    class Database
    {
        private MySqlConnection sqlconn;
        string server; //"46.101.139.249";
        string password; //"Qwerty12344321+";
        string login; // "serverUser";
        string db; //"testitDB";

        public Database(string server, string login, string pass, string db)
        {
            this.server = server;
            this.login = login;
            this.password = pass;
            this.db = db;
            connectToDb();
        }
        public bool connectToDb()
        {
            bool succes = false;
            // sqlconn = new MySqlConnection($"server = {server}; uid = {login}; pwd= {password}; database= {db};");
            sqlconn = new MySqlConnection("server = " + server + "; uid = " + login + "; pwd= " + password + "; database= " + db + ";");
            succes = true;
            return succes;
        }
        public List<string> ShowTables()
        {
            MySqlTransaction tr = null;
            MySqlDataReader rdr = null;
            sqlconn.Open();
            tr = sqlconn.BeginTransaction();
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = sqlconn;
            cmd.Transaction = tr;
            cmd.CommandText = $"show tables";
            rdr = cmd.ExecuteReader();
            List<string> tables = new List<string>();
            while (rdr.Read())
            {
                tables.Add(rdr.GetString(0));
            }
            rdr.Close();
            tr.Commit();
            sqlconn.Close();
            sqlconn.Dispose();
            return tables;
        }
        public DataTable SelectTable(string tableName)
        {
            MySqlTransaction tr = null;
            MySqlDataReader rdr = null;
            sqlconn.Open();
            tr = sqlconn.BeginTransaction();
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = sqlconn;
            cmd.Transaction = tr;
            cmd.CommandText = $"select * from {tableName};";
            rdr = cmd.ExecuteReader();
            DataTable dt = new DataTable();
            dt.Load(rdr);
            rdr.Close();
            tr.Commit();
            sqlconn.Close();
            sqlconn.Dispose();
            return dt;
        }
        public bool CreateTableAndInsertData(string nameTable, List<string> columnName, List<List<string>> columnData)
        {
            MySqlTransaction tr = null;
            sqlconn.Open();
            tr = sqlconn.BeginTransaction();
            MySqlCommand cmd = new MySqlCommand();
            cmd.Connection = sqlconn;
            cmd.Transaction = tr;
            string columnsStr = "";
            string columnStrwithoutType = "";
            for (int i = 0; i < columnName.Count; i++)
            {
                string clearStr = new string(columnName[i].Where(c => Char.IsLetter(c)).ToArray());
                if (i != columnName.Count - 1)
                {
                    columnsStr += $" {clearStr} varchar(100),";
                    columnStrwithoutType += $" {clearStr},";
                }
                else // last
                {
                    columnsStr += $" {clearStr} varchar(100)";
                    columnStrwithoutType += $" {clearStr}";
                }
            }


            cmd.CommandText = $"create table {nameTable} ( {columnsStr} );";
            cmd.ExecuteNonQuery();

            for (int y = 0; y < columnData[0].Count; y++)
            {
                string valuesStr = "";
                for (int i = 0; i < columnData.Count; i++)
                {                
                    if(i != columnName.Count - 1)
                    {
                        valuesStr += $"'{columnData[i][y]}',";
                    }
                    else
                    {
                        valuesStr += $"'{columnData[i][y]}'";
                    }
                    
                }
                cmd.CommandText = $"insert into {nameTable}({columnStrwithoutType}) values({valuesStr});";
                cmd.ExecuteNonQuery();
            }


            sqlconn.Close();
            sqlconn.Dispose();
            return true;
        }
    }
}
