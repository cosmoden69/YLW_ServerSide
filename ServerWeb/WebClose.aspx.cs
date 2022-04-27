using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace YLW_WebService.ServerSide
{
    public partial class WebClose : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm_LoadComplete;
        }
        private void WebForm_LoadComplete(object sender, EventArgs e)
        {
            ClientScript.RegisterStartupScript(typeof(Page), "closePage1", "window.close();", true);
            ClientScript.RegisterStartupScript(typeof(Page), "closePage2", "window.open('','haesung').close();", true);
        }
    }
}