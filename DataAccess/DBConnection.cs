using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Reflection;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
namespace DataAccess
{
    public class DBConnection
    {
        public enum DatabaseType
        {
            AS400,
            SQL,
            AS400N,
            TYESQL,
            ACCESS
        }
        
	
        public DataTable ExcuteQueryString(string strcmd,DatabaseType DBtype)
        {
            DataSet ds = new DataSet();
            if(DBtype == DatabaseType.AS400)
            {
                
                    OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400"));
                    OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(ds);
                    return ds.Tables[0];
            }
            else if (DBtype == DatabaseType.AS400N)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400N"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                return ds.Tables[0];
            }

            else if (DBtype == DatabaseType.ACCESS)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("ACCESS"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                da.Fill(ds);
                return ds.Tables[0];
            }
            else
            {
                    SqlConnection conn = new SqlConnection(GetConnectionString("SQL"));
                    SqlCommand cmd = new SqlCommand(strcmd,conn);
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    return ds.Tables[0];
            }
            
        }

        public DataTable ExecuteQueryExcel(string cmd,string path,string sheet)
        {
            string str; 
          
                str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES'";
            
            OleDbConnection conn = new OleDbConnection(str);
            OleDbDataAdapter da = new OleDbDataAdapter("select * from [" + sheet + "$]", conn);
            DataSet ds = new DataSet();
             da.Fill(ds, "table1");
             return ds.Tables[0];

            //string a = ds.Tables[0].Rows[22][0].ToString();
        }

        public DataTable ExecuteQueryCSV(string cmd , string filename)
        {
           
            string connString = string.Format(
                @"Provider=Microsoft.Jet.OleDb.4.0; Data Source={0};Extended Properties=""Text;HDR=YES;FMT=Delimited""",
                Path.GetDirectoryName(filename)
            );
            DataSet ds = new DataSet("CSV File");
            using (OleDbConnection conn = new OleDbConnection(connString))
            {
                conn.Open();
                string query = "SELECT * FROM [" + Path.GetFileName(filename) + "]";
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn))
                {
                    
                    adapter.Fill(ds,"table1");
                }
            }
            return ds.Tables[0];
        }

        public int ExcuteNonQueryString(string strcmd, DatabaseType DBtype)
        {
            int result = 0;
            DataSet ds = new DataSet();
            if (DBtype == DatabaseType.AS400)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    result = cmd.ExecuteNonQuery();
                }
                catch
                {
                    result = 0;
                }
                finally
                {
                    conn.Close();
                }
            }
            else if (DBtype == DatabaseType.AS400N)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400N"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    result = cmd.ExecuteNonQuery();
                }
                catch
                {
                    result = 0;
                }
                finally
                {
                    conn.Close();
                }
            }
            else
            {
                SqlConnection conn = new SqlConnection(GetConnectionString("SQL"));
                SqlCommand cmd = new SqlCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    result = cmd.ExecuteNonQuery();
                }
                catch
                {
                    result = 0;
                }
                finally
                {
                    conn.Close();
                }
            }
            return result;
        }

        public object ExcuteScalar(string strcmd,DatabaseType DBtype)
        {
            object x = new object();
            if (DBtype == DatabaseType.AS400)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    x = cmd.ExecuteScalar();

                }
                catch
                {
                }
                finally
                {
                    conn.Close();
                }

                return x;
            }
            else if (DBtype == DatabaseType.AS400N)
            {

                OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400N"));
                OleDbCommand cmd = new OleDbCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    x = cmd.ExecuteScalar();

                }
                catch
                {
                }
                finally
                {
                    conn.Close();
                }

                return x;
            }
            else
            {
                SqlConnection conn = new SqlConnection(GetConnectionString("SQL"));
                SqlCommand cmd = new SqlCommand(strcmd, conn);
                try
                {
                    conn.Open();
                    x = cmd.ExecuteScalar();

                }
                catch
                {
                }
                finally
                {
                    conn.Close();
                }

                return x;
            }
           
            
           
        }

        protected static string GetConnectionString(string ConnType)
        {
            if (ConnType == "AS400")
            {
                return ConfigurationManager.ConnectionStrings["AS400"].ConnectionString;
            }
            else if (ConnType == "AS400N")
            {
                return ConfigurationManager.ConnectionStrings["AS400N"].ConnectionString;
            }
            else if (ConnType == "TYESQL")
            {
                return ConfigurationManager.ConnectionStrings["TYESQL"].ConnectionString;
            }
            else
            {
                return ConfigurationManager.ConnectionStrings["SQL"].ConnectionString;
            }
        }
        public bool TransferData(string strcmd)
        {
            bool result;
            OleDbConnection conn = new OleDbConnection(GetConnectionString("AS400"));
            OleDbCommand cmd = new OleDbCommand(strcmd, conn);
            try
            {
                conn.Open();
                cmd.ExecuteNonQuery();
                result = true;
            }
            catch {
                result = false;
            }
            finally
            {
                conn.Close();
            }
            return result;
        }

        public List<DataHeader> ConvertDataTableToListHeader(DataTable dt)
        {
            PropertyInfo[] AllProps;
             
            int i = 0;
            AllProps = new DataHeader().GetType().GetProperties();
            List<DataHeader> lHead = new List<DataHeader>();
            DataHeader dH;
            foreach (DataRow dr in dt.Rows)
            {
                dH = new DataHeader();
                i = 0;
                foreach (PropertyInfo PropA in AllProps)
                {
                    PropA.SetValue(dH, dr[i], null);
                    i++;
                }
                lHead.Add(dH);
            }
            return lHead;
            
        }

        public List<DataDetail> ConvertDataTableToListDetail(DataTable dt)
        {
            PropertyInfo[] AllProps;

            int i = 0;
            AllProps = new DataDetail().GetType().GetProperties();
            List<DataDetail> lDetail = new List<DataDetail>();
            DataDetail dD;
            foreach (DataRow dr in dt.Rows)
            {
                dD = new DataDetail();
                i = 0;
                foreach (PropertyInfo PropA in AllProps)
                {
                    PropA.SetValue(dD, dr[i], null);
                    i++;
                }
                lDetail.Add(dD);
            }
            return lDetail;

        }

        public DataSet ExcuteQuerySP(SqlCommand cmd, DatabaseType DBtype)
        {
            SqlConnection conn = new SqlConnection();
            DataSet ds = new DataSet();
            
            if (DBtype == DatabaseType.TYESQL)
            {
                conn = new SqlConnection(GetConnectionString("TYESQL"));
            }
            else
            {
                conn = new SqlConnection(GetConnectionString("SQL"));
            }
            //SqlCommand cmd = new SqlCommand(strcmd, conn);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Connection = conn;
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            
            da.Fill(ds);
            return ds;

        }

        public int ExcuteNonQuerySP(SqlCommand cmd, DatabaseType DBtype)
        {
            SqlConnection conn = new SqlConnection();
            
            if (DBtype == DatabaseType.TYESQL)
            {
                conn = new SqlConnection(GetConnectionString("TYESQL"));
            }
            else
            {
                conn = new SqlConnection(GetConnectionString("SQL"));
            }
            cmd.Connection = conn;
            //SqlCommand cmd = new SqlCommand(strcmd, conn);
            cmd.CommandType = CommandType.StoredProcedure;
            int result = 0;
            
                conn.Open();
                result = cmd.ExecuteNonQuery();
                conn.Close();
            
           
            return result;
        }

        public DataTable ExcuteQueryExcel(string FilePath, string Extension, string isHDR,string sheet)
        {
            string conStr = "";
            switch (Extension)
            {
                case ".xls": //Excel 97-03
                    conStr = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                    break;
                case ".xlsx": //Excel 07
                    conStr = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                    break;
            }

            conStr = String.Format(conStr, FilePath, isHDR);
            OleDbConnection connExcel = new OleDbConnection(conStr);
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter oda = new OleDbDataAdapter();
            DataTable dt = new DataTable();
            cmdExcel.Connection = connExcel;

            //Get the name of First Sheet
            connExcel.Open();

            DataTable dtExcelSchema;
            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            //Get all sheet name (Not used)
            string SheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            connExcel.Close();
            //Read Data from First Sheet
            if (!sheet.Contains("$"))
            {
                SheetName = sheet + "$";
            }

            else
            {
                SheetName = sheet;
            }
            connExcel.Open();
            cmdExcel.CommandText = "SELECT * From [" + SheetName + "]";
            oda.SelectCommand = cmdExcel;
            oda.Fill(dt);
            connExcel.Close();
            return dt;

        }

        public List<string> ListSheetInExcel(string filePath)
        {
            OleDbConnectionStringBuilder sbConnection = new OleDbConnectionStringBuilder();
            String strExtendedProperties = String.Empty;
            sbConnection.DataSource = filePath;
            if (Path.GetExtension(filePath).Equals(".xls"))//for 97-03 Excel file
            {
                sbConnection.Provider = "Microsoft.Jet.OLEDB.4.0";
                strExtendedProperties = "Excel 8.0;HDR=Yes;IMEX=1";//HDR=ColumnHeader,IMEX=InterMixed
            }
            else if (Path.GetExtension(filePath).Equals(".xlsx"))  //for 2007 Excel file
            {
                sbConnection.Provider = "Microsoft.ACE.OLEDB.12.0";
                strExtendedProperties = "Excel 12.0;HDR=Yes;IMEX=1";
            }
            sbConnection.Add("Extended Properties", strExtendedProperties);
            List<string> listSheet = new List<string>();
            using (OleDbConnection conn = new OleDbConnection(sbConnection.ToString()))
            {
                conn.Open();
                DataTable dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    if (drSheet["TABLE_NAME"].ToString().Contains("$"))//checks whether row contains '_xlnm#_FilterDatabase' or sheet name(i.e. sheet name always ends with $ sign)
                    {
                        listSheet.Add(drSheet["TABLE_NAME"].ToString());
                    }
                }
            }
            return listSheet;
        }

       

    }
}
