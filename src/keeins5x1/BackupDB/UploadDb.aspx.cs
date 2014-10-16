using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BackupDB_UploadDb : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            //uploadDb.Attributes.Add("accept", "application/x-msaccess");
        }
    }

    protected void btnUpload_Click(object sender, EventArgs e)
    {
        if (!uploadDb.HasFile)
        {
            ShowMessage("請選擇資料庫檔案");
        }
        else
        {
            string file_name = uploadDb.FileName;

            file_name = file_name.Substring(file_name.LastIndexOf('.') + 1);

            if (!file_name.Equals("mdb", StringComparison.OrdinalIgnoreCase))
            {
                ShowMessage("僅限定 *.mdb 資料檔案");

                return;
            }

            string file_online = Server.MapPath("../App_Data/INSDB.mdb");

            if (File.Exists(file_online))
            {
                try
                {
                    string dir_bak = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string dir_bak_path = Server.MapPath(String.Format("../App_Data/{0}", dir_bak));

                    string file_bak = Server.MapPath(String.Format("../App_Data/{0}/INSDB.mdb", dir_bak));

                    Directory.CreateDirectory(dir_bak_path);

                    File.Copy(file_online, file_bak, true);

                    StreamWriter sw = File.CreateText(Server.MapPath(String.Format("../App_Data/{0}/Backup.txt", dir_bak)));
                    sw.AutoFlush = true;

                    sw.WriteLine("00000000");
                    sw.WriteLine("99999999");
                    sw.Close();
                }
                catch (Exception)
                {
                    ShowMessage("原始檔案備份失敗，請稍後再試");
                    
                    return;
                }
            }

            uploadDb.SaveAs(file_online);
        }
    }

    protected void ShowMessage(string message)
    {
        Page.ClientScript.RegisterStartupScript(GetType(), Guid.NewGuid().ToString(), String.Format("alert('{0}');", message), true);
    }
}