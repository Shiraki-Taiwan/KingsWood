using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.Security;

public partial class Login_Login : System.Web.UI.Page
{
	protected void Page_Load(object sender, EventArgs e)
	{
		if (!IsPostBack)
		{
			if (!string.IsNullOrEmpty(Request.Form["ID"]))
			{
				if (Request.Form["ID"].Equals("ins", StringComparison.OrdinalIgnoreCase))
				{
					Session["GroupType"] = Guid.NewGuid().ToString();

					FormsAuthentication.Authenticate("ins", "ins");
					//FormsAuthentication.RedirectFromLoginPage("ins", true);
					//FormsAuthentication.
					//FormsAuthentication.SetAuthCookie("ins", true);
				}				
			}
		}

		Response.Write("User.Identity.IsAuthenticated = " + User.Identity.IsAuthenticated + "<br />");

		if (Session["GroupType"] == null)
			Response.Write(Session["GroupType"] == null);
		else
			Response.Write(Session["GroupType"].ToString());
	}
}