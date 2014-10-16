using System;
using System.Collections.Generic;
using System.IO;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BackupDB_BackupINSDB : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
	{
		if (!IsPostBack)
		{
			string file_online = Server.MapPath("../App_Data/INSDB.mdb");

			Response.Clear();

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

					Response.Write("1");
				}
				catch (Exception)
				{
					Response.Write("0");
				}
			}

			Response.End();
		}
    }
}