using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Text;

namespace YLW_WebService.ServerSide
{
    public class webCommon
    {
        public static string GetScript(string url)
        {
            string script = "var winPop; ";
            script += " if (!winPop || (winPop && winPop.closed)) ";
            script += " { ";
            script += "     winPop = window.open('" + url + "', 'haesung', 'left=1024,top=1024,width=10px,height=10px,location=no,toolbar=no,menubar=no,scrollbars=no,resizable=no,visible=none'); ";
            script += " } ";
            script += " else ";
            script += " { ";
            script += "     winPop.location.href = '" + url + "' ; ";
            script += " }; ";
            //script += " winPop.setTimeout(() => { winPop.close(); }, 2000); ";
            script += " winPop.blur(); ";
            script += " window.focus(); ";

            return script;
        }

        public static void WebFormPageLoad(ClientScriptManager ClientScript, string para)
        {
            string script = webCommon.GetScript(para);
            ClientScript.RegisterStartupScript(typeof(Page), "run", "<script language=javascript>" + script + "</script>");
        }

        public static void WebFormLoadComplete(ClientScriptManager ClientScript)
        {
            ClientScript.RegisterStartupScript(typeof(Page), "closePage1", "window.close();", true);
            //ClientScript.RegisterStartupScript(typeof(Page), "closePage2", "window.open('','haesung').close();", true);
        }

        public static string WebFormRequest(string url)
        {
            string responseText = string.Empty;
            HttpWebRequest request = HttpWebRequest.Create(url) as HttpWebRequest;
            request.Method = "GET";
            request.Timeout = 30 * 1000; // 30초
            request.Headers.Add("Access-Control-Allow-Origin", "*"); // 헤더 추가 방법

            using (HttpWebResponse resp = (HttpWebResponse)request.GetResponse())
            {
                HttpStatusCode status = resp.StatusCode;
                Console.WriteLine(status);  // 정상이면 "OK"

                Stream respStream = resp.GetResponseStream();
                using (StreamReader sr = new StreamReader(respStream))
                {
                    responseText = sr.ReadToEnd();
                }
            }
            return responseText;
        }
    }
}