using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace YLW_WebService.ServerSide
{
    public partial class wfOpenHRFileView : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm_LoadComplete;

            if (!IsPostBack)
            {
                string value = Request.QueryString["para"];
                webCommon.WebFormPageLoad(ClientScript, "http://localhost:8080/OpenHRFileView/" + value + "");
            }
        }
        private void WebForm_LoadComplete(object sender, EventArgs e)
        {
            webCommon.WebFormLoadComplete(ClientScript);
        }
    }
}