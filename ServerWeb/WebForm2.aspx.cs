using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace YLW_WebService.ServerSide
{
    public partial class WebForm2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm_LoadComplete;

            if (!IsPostBack)
            {
                string value = Request.QueryString["para"];
                //ClientScript.RegisterStartupScript(typeof(Page), "popup", "<script language=javascript>window.open('WebForm3.aspx?para=" + value + "','','width=10px,height=10px')</script>");
                //string script = " window.open('http://localhost:8080/OpenDocx/" + value + "','','width=10px,height=10px').close(); ";
                //ClientScript.RegisterStartupScript(typeof(Page), "popup", script, true);
                //ClientScript.RegisterStartupScript(typeof(Page), "closePage", "window.close();", true);

                webCommon.WebFormPageLoad(ClientScript, "http://localhost:8080/OpenDocx/" + value + "");

            }
        }
        private void WebForm_LoadComplete(object sender, EventArgs e)
        {
            webCommon.WebFormLoadComplete(ClientScript);
        }
    }
}