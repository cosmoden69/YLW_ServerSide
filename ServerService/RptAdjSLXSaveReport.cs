using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.Schema;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using YLWService;
using YLWService.Extensions;

namespace YLW_WebService.ServerSide
{
    public class RptAdjSLXSaveReport
    {
        public string myPath = Application.StartupPath;

        public RptAdjSLXSaveReport()
        {
        }

        public RptAdjSLXSaveReport(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            return RptMain(para, ref rptPath, ref rptName);
        }

        public Response RptMain(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersCS";
                security.methodId = "ReportQuery";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock10");
                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");
                dt.Columns.Add("ReportType");

                dt.Clear();
                DataRow dr = dt.Rows.Add();
                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;
                dr["ReportType"] = para.ReportType;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null || yds.Tables.Count < 1 || yds.Tables[0].Rows.Count < 1)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                string fileName = yds.Tables[0].Rows[0]["FileName"] + "";
                string fileSeq = yds.Tables[0].Rows[0]["FileSeq"] + "";

                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string fileBase64 = YLWService.MTRServiceModule.CallMTRFileDownloadBase64(security, fileSeq, "0", "0");
                if (fileBase64 == "")
                {
                    return new Response() { Result = -1, Message = "보고서파일을 생성할 수 없습니다" };
                }
                byte[] bytes_file = Convert.FromBase64String(fileBase64);
                FileStream orgFile = new FileStream(sSample1Relt, FileMode.Create);
                orgFile.Write(bytes_file, 0, bytes_file.Length);
                orgFile.Close();

                rptName = fileName;
                rptPath = sSample1Relt;

                return new Response() { Result = 1, Message = "OK" };
            }
            catch (Exception ex)
            {
                return new Response() { Result = -99, Message = ex.Message };
            }
        }

        public Response RptHistoryPost(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersCS";
                security.methodId = "ReportHistoryQuery";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock10");
                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");
                dt.Columns.Add("ReportType");
                dt.Columns.Add("Seq");

                dt.Clear();
                DataRow dr = dt.Rows.Add();
                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;
                dr["ReportType"] = para.ReportType;
                dr["Seq"] = para.Seq;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null || yds.Tables.Count < 1 || yds.Tables[0].Rows.Count < 1)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                string fileName = yds.Tables[0].Rows[0]["FileName"] + "";
                string fileSeq = yds.Tables[0].Rows[0]["FileSeq"] + "";

                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string fileBase64 = YLWService.MTRServiceModule.CallMTRFileDownloadBase64(security, fileSeq, "0", "0");
                if (fileBase64 == "")
                {
                    return new Response() { Result = -1, Message = "보고서파일을 생성할 수 없습니다" };
                }
                byte[] bytes_file = Convert.FromBase64String(fileBase64);
                FileStream orgFile = new FileStream(sSample1Relt, FileMode.Create);
                orgFile.Write(bytes_file, 0, bytes_file.Length);
                orgFile.Close();

                rptName = fileName;
                rptPath = sSample1Relt;

                return new Response() { Result = 1, Message = "OK" };
            }
            catch (Exception ex)
            {
                return new Response() { Result = -99, Message = ex.Message };
            }
        }
    }
}