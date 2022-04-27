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
    public partial class WebTest1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebFormIn_LoadComplete;

            string value = "{\"PgmName\":\"FrmAdjHREmpInfoMgmt\",\"CompanySeq\":\"1\",\"UserID\":\"msuser@metrosoft.co.kr\"}";
            string url = "http://localhost:8089/OpenForm/" + value + "";
            //string script = "";
            //script += " var x = new XMLHttpRequest(); ";
            //script += " x.open('GET', '" + url + "', false); ";
            //script += " x.withCredentials = true; ";
            //script += " x.onload = function() { alert(x.responseText); }; ";
            //script += " x.send(); ";
            //ClientScript.RegisterStartupScript(typeof(Page), "run", "<script language=javascript>" + script + "</script>");

            //string url = "http://localhost:8089/OpenPost";
            //string script = "";
            //script += " var x = new XMLHttpRequest(); ";
            //script += " x.open('POST', '" + url + "', true); ";
            //script += " x.setRequestHeader('method', 'access-control-request-method'); ";
            //script += " x.setRequestHeader('Content-type', 'text/plain'); ";
            ////script += " x.send(JSON.stringify(" + value + ")); ";
            ////script += "const params = {email: '123', password: '456' };";
            ////script += " x.send(JSON.stringify(params)); ";
            //script += " x.onload = function() { alert(x.responseText); }; ";
            ////script += " x.send('111'); ";
            //script += " x.send(JSON.stringify(value='123')); ";
            //ClientScript.RegisterStartupScript(typeof(Page), "run", "<script language=javascript>" + script + "</script>");

            //string responseText = string.Empty;
            //HttpWebRequest request = HttpWebRequest.Create(url) as HttpWebRequest;
            //request.Method = "GET";
            ////request.Timeout = 30 * 1000; // 30초
            ////request.Headers.Add("Access - Control - Allow - Origin", "http://");
            //request.Headers.Add("Authorization", "BASIC SGVsbG8="); // 헤더 추가 방법

            //using (HttpWebResponse resp = (HttpWebResponse)request.GetResponse())
            //{
            //    HttpStatusCode status = resp.StatusCode;
            //    Console.WriteLine(status);  // 정상이면 "OK"

            //    Stream respStream = resp.GetResponseStream();
            //    using (StreamReader sr = new StreamReader(respStream))
            //    {
            //        responseText = sr.ReadToEnd();
            //    }
            //    JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
            //    responseText = JsonConvert.DeserializeObject<string>(responseText, settings);
            //    if (responseText == "OK") Console.WriteLine(responseText);
            //}

            //string script = " var winPop = window.open('" + url + "', 'haesung', 'left=1024,top=1024,width=10px,height=10px,location=no,toolbar=no,menubar=no,scrollbars=no,resizable=no,status=no'); ";
            ////string script = " var winPop = window.open('" + url + "', 'haesung'); ";
            ////string script = " var winPop = window.open('" + url + "', 'haesung'); ";
            ////script += " winPop.onload = function() { setTimeout('window.close()', 1000); }; ";
            //script += " winPop.onload = function() { alert('111'); }; ";
            ////script += " winPop.addEventListener('load', function() { winPop.close(); }); ";
            ////script += " winPop.onload = function() { open('', '_self').close(); }; ";
            //script = " alert('111'); ";
            //ClientScript.RegisterStartupScript(typeof(Page), "run1", "<script language=javascript>" + script + "</script>");
            ClientScript.RegisterStartupScript(typeof(Page), "run1", "<script language=javascript>WebSvc()</script>");
        }

        private void WebFormIn_LoadComplete(object sender, EventArgs e)
        {
            ClientScript.RegisterStartupScript(typeof(Page), "closePage", "window.close();", true);
        }
    }
}