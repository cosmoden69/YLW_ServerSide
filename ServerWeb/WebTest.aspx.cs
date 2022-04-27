using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace YLW_WebService.ServerSide
{
    public partial class WebTest : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm_LoadComplete;

            if (!IsPostBack)
            {
                string value = "{\"PgmName\":\"FrmAdjHREmpInfoMgmt\",\"CompanySeq\":\"1\",\"UserID\":\"msuser@metrosoft.co.kr\"}";
                ClientScript.RegisterStartupScript(typeof(Page), "popup", "<script language=javascript>window.open('WebTest1.aspx?para=" + value + "','','width=10px,height=10px')</script>");
            }
        }
        private void WebForm_LoadComplete(object sender, EventArgs e)
        {
            webCommon.WebFormLoadComplete(ClientScript);
        }
    }
}