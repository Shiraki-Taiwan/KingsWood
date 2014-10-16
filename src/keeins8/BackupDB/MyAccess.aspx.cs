using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class BackupDB_MyAccess : System.Web.UI.Page
{
	private string ConnStr
	{
		get
		{
			return "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Server.MapPath("../App_Data/HAS.accdb");
		}
	}

    protected void Page_Load(object sender, EventArgs e)
    {

    }

	protected void btnQuery_Click(object sender, EventArgs e)
	{
		if (String.IsNullOrEmpty(txtSQL.Text)) return;

		if (txtSQL.Text.Contains(";"))
		{
			string[] arrSQL = txtSQL.Text.Split(';');

			using (OleDbConnection conn = new OleDbConnection(ConnStr))
			{
				conn.Open();

				using (OleDbCommand cmd = conn.CreateCommand())
				{
					foreach (string sql in arrSQL)
					{
						string t_sql = sql.Trim();

						if (String.IsNullOrEmpty(t_sql)) continue;

						cmd.CommandText = t_sql;

						try
						{
							Response.Write("(" + cmd.ExecuteNonQuery() + ")" + t_sql + "<br />");
						}
						catch (Exception ex)
						{
							Response.Write("(0)" + t_sql + "(" + ex.Message + ")" + "<br />");
						}
					}
				}

				conn.Close();
			}
		}
		else
		{
			using (System.Data.DataTable dt = new System.Data.DataTable())
			{
				using (OleDbConnection conn = new OleDbConnection(ConnStr))
				{
					conn.Open();

					using (OleDbCommand cmd = conn.CreateCommand())
					{
						cmd.CommandText = txtSQL.Text;

						using (OleDbDataAdapter oda = new OleDbDataAdapter(cmd))
						{
							oda.Fill(dt);
						}
					}

					conn.Close();
				}

				gvResult.DataSource = dt;
				gvResult.DataBind();
			}
		}
	}
}