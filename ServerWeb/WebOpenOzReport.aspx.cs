using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
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

using YLWService;
using YLWService.Extensions;

namespace YLW_WebService.ServerSide
{
    public partial class WebOpenOzReport : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            this.LoadComplete += WebForm_LoadComplete;

            string value = Request.QueryString["para"];
            JObject json = JObject.Parse(value);
            int companySeq = Convert.ToInt32(json["CompanySeq"]);
            string rptname = json["ReportName"].ToString();
            string ParamStr = json["ParamStr"].ToString();
            JObject param1 = JObject.Parse(ParamStr);
            string parameterSeq = param1["ParameterSeq"].ToString();
            string xmlString = GetXmlDataString(companySeq, parameterSeq);

            string url = "http://ksystem.metro070.com:8200";

            if (!IsPostBack)
            {
                string winHTML = ""
                    + " <%@ Page Language='C#' AutoEventWireup='true' CodeBehind='WebOpenOzReport1.aspx.cs' Inherits='YLW_WebService.ServerSide.WebOpenOzReport1' %> " + Environment.NewLine
                    + " <!DOCTYPE html>                                                                                                          " + Environment.NewLine
                    + " <html xmlns='http://www.w3.org/1999/xhtml'>                                                                              " + Environment.NewLine
                    + " <head>                                                                                                                   " + Environment.NewLine
                    + " <title> OZReport </title>                                                                                                " + Environment.NewLine
                    + " <meta http - equiv = 'X-UA-Compatible' content = 'IE=edge' />                                                            " + Environment.NewLine
                    + " <!--20210914.tjjang, ACE인 경우에는 / Oz70 를 넣어야 하고, EVER 인 경우 / Oz70 을 뺴야한다-->                            " + Environment.NewLine
                    + " <script type = 'text/javascript' src = '/Oz70/ozhviewer/jquery-1.8.3.min.js' ></script>                                  " + Environment.NewLine
                    + " <link rel = 'stylesheet' href = '/Oz70/ozhviewer/jquery-ui.css' type = 'text/css' />                                     " + Environment.NewLine
                    + " <script type = 'text/javascript' src = '/Oz70/ozhviewer/jquery-ui.min.js' ></script>                                     " + Environment.NewLine
                    + " <link rel = 'stylesheet' href = '/Oz70/ozhviewer/ui.dynatree.css?WEBK20190919=12' type = 'text/css' />                   " + Environment.NewLine
                    + " <script type = 'text/javascript' src = '/Oz70/ozhviewer/jquery.dynatree.js?WEBK20190919=12' charset = 'utf-8' ></script> " + Environment.NewLine
                    + " <script type = 'text/javascript' src = '/Oz70/ozhviewer/OZJSViewer.js?WEBK20190919=12' charset = 'utf-8' ></script>      " + Environment.NewLine
                    + " </head>                                                                                                                  " + Environment.NewLine
                    + " <body style = 'height:100%' >                                                                                            " + Environment.NewLine
                    + " <div id = 'OZViewer' style = 'width:100%; height:100%;' ></div>                                                          " + Environment.NewLine
                    + " <script type = 'text/javascript' >                                                                                       " + Environment.NewLine
                    + "     function SetOZParamters_OZViewer(){                                                                                  " + Environment.NewLine
                    + "         var oz;                                                                                                          " + Environment.NewLine
                    + "         var strOdiFileName = '" + rptname + "';                                                                          " + Environment.NewLine
                    + "         var strSiteFileName = '" + rptname + "';                                                                         " + Environment.NewLine
                    + "         xmlString = 'xmlData=<?xml version=\"1.0\" encoding=\"UTF-8\"?>" + xmlString + "'; " + Environment.NewLine
                    + "         var xmlString;                                                                                                   " + Environment.NewLine
                    + "         oz = document.getElementById('OZViewer');                                                                        " + Environment.NewLine
                    //+ "         //Default -> 레포트설정에 따라 변경가능                                                                        " + Environment.NewLine
                    + "         oz.sendToActionScript('viewer.isframe', 'false');                                                                " + Environment.NewLine
                    //+ "         //print미리보기 제거                                                                                           " + Environment.NewLine
                    + "         oz.sendToActionScript('print.mode', 'false');                                                                    " + Environment.NewLine
                    //+ "         //최소 글꼴 크기 조정                                                                                          " + Environment.NewLine
                    + "         oz.sendToActionScript('viewer.fontdpi', 'auto');                                                                 " + Environment.NewLine
                    //+ "         //#20200806 서브폼 출력물의 경우 100%를 기본으로 설정                                                          " + Environment.NewLine
                    + "         oz.sendToActionScript('viewer.zoom', '100');                                                                     " + Environment.NewLine
                    + "         oz.sendToActionScript('odi.odinames', strOdiFileName);                                                           " + Environment.NewLine
                    + "         oz.sendToActionScript('odi.' + strOdiFileName + '.usescheduleddata', 'ozp:///sdmmaker_html5(xml)/sdmmaker_ylw_h.js'); " + Environment.NewLine
                    + "         oz.sendToActionScript('odi.' + strOdiFileName + '.pcount', '1');                                                 " + Environment.NewLine
                    + "         oz.sendToActionScript('odi.' + strOdiFileName + '.args1', xmlString);                                            " + Environment.NewLine
                    + "         oz.sendToActionScript('connection.servlet', '/Oz70/');                                                           " + Environment.NewLine
                    + "         oz.sendToActionScript('connection.reportname', '/' + strSiteFileName + '.ozr');                                  " + Environment.NewLine
                    + "         oz.sendToActionScript('connection.displayname', strOdiFileName + '_pdf');                                        " + Environment.NewLine
                    + "         oz.sendToActionScript('connection.formfromserver', 'true');                                                      " + Environment.NewLine
                    //+ "         //폰트 파라미터                                                                                                " + Environment.NewLine
                    + "         oz.sendToActionScript('pdf.fontembedding', 'true');                                                              " + Environment.NewLine
                    + "         oz.sendToActionScript('information.debug', 'true');                                                              " + Environment.NewLine
                    //+ "         //SDM FILE 에러 수정적용                                                                                       " + Environment.NewLine
                    + "         oz.sendToActionScript('connection.datafromserver', 'false');                                                     " + Environment.NewLine
                    + "         return true;                                                                                                     " + Environment.NewLine
                    + "     };                                                                                                                   " + Environment.NewLine
                    + "     var opt = [];                                                                                                        " + Environment.NewLine
                    + "     opt['print_exportfrom'] = 'scheduler';                                                                               " + Environment.NewLine
                    + "     start_ozjs('OZViewer', '/Oz70/ozhviewer/', opt);                                                                     " + Environment.NewLine
                    + " </script>                                                                                                                " + Environment.NewLine
                    + " </body>                                                                                                                  " + Environment.NewLine
                    + " </html>                                                                                                                  " + Environment.NewLine;

                string mypath = HttpContext.Current.Server.MapPath("~/Temp");
                string myfile = Guid.NewGuid().ToString() + ".aspx";
                string aspxtempfile = mypath + @"\" + myfile;
                // Write the string to a file.
                System.IO.StreamWriter file = new System.IO.StreamWriter(aspxtempfile, false, Encoding.UTF8);
                file.WriteLine(winHTML);
                file.Close();

                string script = "";
                script += " var popupWidth = 1000; ";
                script += " var popupHeight = window.screen.height; ";
                script += " var popupX = (window.screen.width / 2) - (popupWidth / 2); ";    // 만들 팝업창 width 크기의 1/2 만큼 보정값으로 빼주었음
                script += " var popupY = (window.screen.height / 2) - (popupHeight / 2); ";  // 만들 팝업창 height 크기의 1/2 만큼 보정값으로 빼주었음";
                script += " var win = window.open('Temp/" + myfile + "?para1=" + value + "','','status=no, height=' + popupHeight + ', width=' + popupWidth + ', left=' + popupX + ', top=' + popupY);";
                ClientScript.RegisterStartupScript(typeof(Page), "popup", "<script language=javascript>" + script  + "</script>");
            }
        }
        private void WebForm_LoadComplete(object sender, EventArgs e)
        {
            ClientScript.RegisterStartupScript(typeof(Page), "closePage", "window.close();", true);
        }

        private string GetXmlDataString(int companySeq, string parameterSeq)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisAdjSLRptParameters";
                security.methodId = "Query";
                security.companySeq = companySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock1");
                dt.Columns.Add("ParameterSeq");
                dt.Clear();
                DataRow dr = dt.Rows.Add();
                dr["ParameterSeq"] = parameterSeq;

                DataSet yds = YLWService.MTRServiceModule.CallMTRServiceCallPost(security, ds);
                if (yds == null) return null;
                //string xml = yds.Tables["DataBlock1"].Rows[0]["Params"] + "";

                string json = yds.Tables["DataBlock1"].Rows[0]["Params"] + "";
                JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
                DataSet pds = JsonConvert.DeserializeObject<DataSet>(json, settings);

                //DataTable dtB = pds.Tables["DataBlock13"];
                //if (dtB != null && dtB.Rows.Count > 0)
                //{
                //    string fileSeq = Utils.ConvertToString(dtB.Rows[0]["CostRcptFileSeq"]);
                //    if (Utils.ToInt(fileSeq) != 0)
                //    {
                //        DataSet pds1 = YLWService.MTRServiceModule.CallMTRFileDownload(security, fileSeq, "", "");
                //        if (pds1 != null && pds1.Tables.Count > 0)
                //        {
                //            DataTable dtB1 = pds1.Tables[0];
                //            if (dtB1 != null && dtB1.Rows.Count > 0)
                //            {
                //                sSampleDocx = myPath + @"\보고서\출력설계_1715_서식_인보이스(외부용)_Image.docx";
                //                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                //                sRet = SetSample_Image(sSampleDocx, sSampleXSD, pds1, sSampleAddFile);
                //                if (sRet != "")
                //                {
                //                    return new Response() { Result = -1, Message = sRet };
                //                }
                //                addFiles.Add(sSampleAddFile);
                //            }
                //        }
                //    }
                //}
                //xml = xml.Replace("&", "&amp;");
                //xml = xml.Replace("\'", "&apos;");
                //xml = xml.Replace("\"", "&quot;");
                //xml = xml.Replace("\r", "&#xD;");
                //xml = xml.Replace("\n", "&#xA;");
                //xml = xml.Replace("\t", "&#x9;");
                return objectToXml(pds);
            }
            catch (Exception ex)
            {
                return "";
            }
        }

        private string objectToXml(DataSet ds)
        {
            var xml = "";
            xml += "<ROOT>";

            try
            {
                for (var t = 0; t < ds.Tables.Count; t++)
                {
                    var dt = ds.Tables[t];
                    var TableNm = dt.TableName;

                    for (var r = 0; r < dt.Rows.Count; r++)
                    {
                        xml += "<" + TableNm + ">";
                        for (var c = 0; c < dt.Columns.Count; c++)
                        {
                            var dr = dt.Rows[r];
                            var colNm = dt.Columns[c];
                            string rowVal = Utils.ConvertToString(dt.Rows[r][c]);

                            if ((rowVal.ToString()).IndexOf("&") != -1) rowVal = rowVal.Replace("&", "&amp;");
                            if ((rowVal.ToString()).IndexOf("<") != -1) rowVal = rowVal.Replace("<", "&lt;");
                            if ((rowVal.ToString()).IndexOf(">") != -1) rowVal = rowVal.Replace(">", "&gt;");
                            if ((rowVal.ToString()).IndexOf("\'") != -1) rowVal = rowVal.Replace("\'", "&apos;");
                            if ((rowVal.ToString()).IndexOf("\"") != -1) rowVal = rowVal.Replace("\"", "&quot;");
                            if ((rowVal.ToString()).IndexOf("\r") != -1) rowVal = rowVal.Replace("\r", "&#xD;");
                            if ((rowVal.ToString()).IndexOf("\n") != -1) rowVal = rowVal.Replace("\n", "&#xA;");
                            if ((rowVal.ToString()).IndexOf("\t") != -1) rowVal = rowVal.Replace("\t", "&#x9;");
                            if ((rowVal.ToString()).IndexOf("\\") != -1) rowVal = rowVal.Replace("\\", "&#92;");

                            xml += "<" + colNm + ">" + rowVal + "</" + colNm + ">";
                        }
                        xml += "</" + TableNm + ">";
                    }
                }
            }
            catch (Exception ex) { throw ex; }

            xml += "</ROOT>";
            return xml;
        }
    }
}