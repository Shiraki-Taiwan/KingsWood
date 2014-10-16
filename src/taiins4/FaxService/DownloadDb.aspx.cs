using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BackupDB_DownloadDb : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //string file_online = Server.MapPath("../App_Data/INSDB.mdb" + File.ReadAllLines("")[0]);
            string file_online = Server.MapPath("../App_Data/" + File.ReadAllLines(Server.MapPath("../App_Data/connection.txt"))[0]);

            DirectoryInfo dir = new DirectoryInfo(Server.MapPath("../App_Data/"));
            DateTime lastDate = DateTime.MinValue;
            
            foreach (FileInfo fi in dir.GetFiles("INSDB*.mdb"))
            {
                if (lastDate <= fi.LastWriteTime)
                {
                    lastDate = fi.LastWriteTime;
                    file_online = fi.FullName;
                }
            }

            //string file_online = Server.MapPath("../App_Data/INSDB.mdb");
            string out_file = "INSDB.mdb";

            try
            {
                FileInfo xpath_file = new FileInfo(file_online);  //要 using System.IO;
                // 將傳入的檔名以 FileInfo 來進行解析（只以字串無法做）
                System.Web.HttpContext.Current.Response.Clear(); //清除buffer
                System.Web.HttpContext.Current.Response.ClearHeaders(); //清除 buffer 表頭
                System.Web.HttpContext.Current.Response.Buffer = false;
                System.Web.HttpContext.Current.Response.ContentType = "application/octet-stream";
                // 檔案類型還有下列幾種"application/pdf"、"application/vnd.ms-excel"、"text/xml"、"text/HTML"、"image/JPEG"、"image/GIF"
                System.Web.HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment;filename=" + System.Web.HttpUtility.UrlEncode(out_file, System.Text.Encoding.UTF8));
                // 考慮 utf-8 檔名問題，以 out_file 設定另存的檔名
                System.Web.HttpContext.Current.Response.AppendHeader("Content-Length", xpath_file.Length.ToString()); //表頭加入檔案大小
                System.Web.HttpContext.Current.Response.WriteFile(xpath_file.FullName);

                // 將檔案輸出
                System.Web.HttpContext.Current.Response.Flush();
                // 強制 Flush buffer 內容
                System.Web.HttpContext.Current.Response.End();

            }
            catch (Exception)
            {

            }
        }
    }
}