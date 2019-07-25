using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Text.RegularExpressions;

namespace Excel2vCard
{

    public class ImprotExcel
    {
       
        private static bool CheckFileExt(string fileExtension, string[] allowUpload)
        {
            bool flag = false;
            for (int i = 0; i < allowUpload.Length; i++)
            {
                if (fileExtension == allowUpload[i])
                {
                    flag = true;
                }
            }
            return flag;
        }
        /// <summary>
        /// 读取Excel文件的内容返回表格数据
        /// </summary>
        /// <param name="FileUrl">文件所在的硬盘路径</param>
        /// <returns></returns>
        public static DataTable ExecletoDt(string FileUrl)
        {
            string strConn = "";
            OleDbConnection conn;
            try
            {
                strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + FileUrl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
                conn = new OleDbConnection(strConn);
                conn.Open();
            }
            catch (Exception ex)
            {
                //LogHelper.Error("ImprotExcel.ExecletoDt", ex);
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + FileUrl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
                conn = new OleDbConnection(strConn);
                conn.Open();
            }
            //默认只读取第一个
            DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            string SheetName = dtSheetName.Rows[0]["TABLE_NAME"].ToString();
            DataSet ds = new DataSet();
            OleDbDataAdapter odda = new OleDbDataAdapter("select * from [" + SheetName + "]", conn);
            odda.Fill(ds);
            return ds.Tables[0];
        }
        /// <summary>
        ///  读取Excel文件的内容返回表格数据
        /// </summary>
        /// <param name="FileUrl">文件所在的硬盘路径</param>
        /// <param name="SortIndex">指定第几个表</param>
        /// <returns></returns>
        public static DataTable ExecletoDt(string FileUrl, int SortIndex)
        {
            string strConn = "";
            OleDbConnection conn;
            try
            {
                strConn = "Provider=Microsoft.Jet.OleDb.4.0;" + "data source=" + FileUrl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
            }
            catch (Exception ex)
            {
                //LogHelper.Error("ImprotExcel.ExecletoDt", ex);

                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "data source=" + FileUrl + ";Extended Properties='Excel 8.0; HDR=YES; IMEX=1'";
            }
            conn = new OleDbConnection(strConn);
            conn.Open();
            //默认只读取第一个
            DataTable dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
            string SheetName = "";
            if (SortIndex > dtSheetName.Rows.Count)
            {
                return null;
            }
            else
            {
                SheetName = dtSheetName.Rows[SortIndex]["TABLE_NAME"].ToString();
                DataSet ds = new DataSet();
                OleDbDataAdapter odda = new OleDbDataAdapter("select * from [" + SheetName + "]", conn);
                odda.Fill(ds);
                return ds.Tables[0];
            }
        }
        /// <summary>
        /// 检测当前行是否是空白行
        /// </summary>
        /// <param name="dr">行</param>
        /// <returns>true不是空白行</returns>
        public static bool CheckIsBlank(DataRow dr)
        {
            bool result = false;
            foreach (object obj in dr.ItemArray)
            {
                if (obj.ToString() != "")
                {
                    result = true;
                    break;
                }
            }
            return result;
        }

        public static DataTable CSVtoSteam(string FileUrl, bool ReadHead)
        {
            int intColCount = 0;
            DataTable mydt = new DataTable();
            DataColumn mydc;
            DataRow mydr;
            string strline;
            string[] aryline;
            SortedList sl = new SortedList();
            System.IO.StreamReader mysr = new System.IO.StreamReader(FileUrl);//cvs文件路径
            while ((strline = mysr.ReadLine()) != null)
            {
                strline = strline.Replace("\"\"", "'");
                MatchCollection col = Regex.Matches(strline, ",\"([^\"]+)\",", RegexOptions.ExplicitCapture);
                IEnumerator ie = col.GetEnumerator();
                while (ie.MoveNext())
                {
                    string patn = ie.Current.ToString();
                    int key = strline.Substring(0, strline.IndexOf(patn)).Split(',').Length;
                    if (!sl.ContainsKey(key))
                    {
                        sl.Add(key, patn.Trim(new char[] { ',', '"' }).Replace("'", "\""));
                        strline = strline.Replace(patn, ",,");
                    }
                }
                aryline = strline.Split(new char[] { ',' });
                for (int i = 0; i < aryline.Length; i++)
                {
                    if (!sl.ContainsKey(i))
                    {
                        sl.Add(i, aryline[i]);
                    }
                }
                if (ReadHead == true)
                {
                    ReadHead = false;
                    intColCount = sl.Count;
                    for (int j = 0; j < intColCount; j++)
                    {
                        mydc = new DataColumn(sl[j].ToString());
                        mydt.Columns.Add(mydc);
                    }
                }
                else
                {
                    mydr = mydt.NewRow();
                    for (int m = 0; m < intColCount; m++)
                    {
                        mydr[m] = sl[m];
                    }
                    mydt.Rows.Add(mydr);
                }
                sl.Clear();
                strline = string.Empty;
            }
            return mydt;
        }
    }
}