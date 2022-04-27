using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace YLW_WebService.ServerSide
{
    public partial class WebForm3 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm3_LoadComplete;

            string value = Request.QueryString["para"];
            string script = "";
            script += " var x = new XMLHttpRequest(); ";
            script += " x.open('GET', 'http://localhost:8080/OpenDocx/" + value + "', false); ";
            script += " x.withCredentials = true; ";
            script += " x.send(); ";
            ClientScript.RegisterStartupScript(typeof(Page), "run", "<script language=javascript>" + script + "</script>");
        }

        private void WebForm3_LoadComplete(object sender, EventArgs e)
        {
            ClientScript.RegisterStartupScript(typeof(Page), "closePage", "window.close();", true);
        }
    }
}