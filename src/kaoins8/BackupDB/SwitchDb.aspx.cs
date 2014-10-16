using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;

public partial class BackupDB_SwitchDb : System.Web.UI.Page
{
    public static readonly object lockObj = new object();

    protected void Page_Load(object sender, EventArgs e)
    {
        lock (lockObj)
        {
            if (!IsPostBack)
            {
                Response.Clear();

                string dbName = Request.Form["shiraki_kingswood_db_name"];

                if (String.IsNullOrEmpty(dbName))
                {
                    Response.Write("0");
                }
                else
                {
                    File.WriteAllText(Server.MapPath("../App_Data/connection.txt"), dbName);

                    Response.Write("1");
                }

                Response.End();
            }
        }
    }
}