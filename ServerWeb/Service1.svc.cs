using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.ServiceModel;
using System.ServiceModel.Activation;
using System.ServiceModel.Web;
using System.Text;
using System.Reflection;
using System.Threading.Tasks;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using YLWService;

namespace YLW_WebService.ServerSide
{
    [AspNetCompatibilityRequirements(RequirementsMode = AspNetCompatibilityRequirementsMode.Allowed)]
    // 참고: "리팩터링" 메뉴에서 "이름 바꾸기" 명령을 사용하여 코드, svc 및 config 파일에서 클래스 이름 "Service1"을 변경할 수 있습니다.
    // 참고: 이 서비스를 테스트하기 위해 WCF 테스트 클라이언트를 시작하려면 솔루션 탐색기에서 Service1.svc나 Service1.svc.cs를 선택하고 디버깅을 시작하십시오.
    public class Service1 : IService1
    {
        public DataSet ServiceCall(YlwSecurityJson security, DataSet ds)
        {
            try
            {
                YlwDataSet yds = YLWService.YLWServiceModule.CallYlwServiceCall(security, ds);
                DataSet result = YLWServiceModule.YDsToDataSet(yds);
                return result;
            }
            catch (Exception ex)
            {
                DataSet dsr = new DataSet();
                DataTable dtr = dsr.Tables.Add("ErrorMessage");
                dtr.Columns.Add("Status");
                dtr.Columns.Add("Message");
                DataRow dr = dtr.Rows.Add();
                dr["Status"] = "ERR";
                dr["Message"] = ex.Message;
                return dsr;
            }
        }

        public DataSet ServiceCallPost(YlwSecurityJson security, DataSet ds)
        {
            try
            {
                return YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
            }
            catch (Exception ex)
            {
                DataSet dsr = new DataSet();
                DataTable dtr = dsr.Tables.Add("ErrorMessage");
                dtr.Columns.Add("Status");
                dtr.Columns.Add("Message");
                DataRow dr = dtr.Rows.Add();
                dr["Status"] = "ERR";
                dr["Message"] = ex.Message;
                return dsr;
            }
        }

        public DataSet GetDataSetPost(int companyseq, string streamdata)
        {
            try
            {
                DataTable dtr = YLWService.YLWServiceModule.GetYlwServiceDataTable(companyseq, streamdata);
                DataSet dsr = new DataSet("Result");
                if (dtr != null) dsr.Tables.Add(dtr);
                return dsr;
            }
            catch (Exception ex)
            {
                DataSet dsr = new DataSet();
                DataTable dtr = dsr.Tables.Add("ErrorMessage");
                dtr.Columns.Add("Status");
                dtr.Columns.Add("Message");
                DataRow dr = dtr.Rows.Add();
                dr["Status"] = "ERR";
                dr["Message"] = ex.Message;
                return dsr;
            }
        }

        public DataSet Fileupload(YlwSecurityJson security, DataSet ds)
        {
            return YLWService.YLWServiceModule.Fileupload(security, ds);
        }

        public string FileuploadGetSeq(YlwSecurityJson security, DataSet ds)
        {
            return YLWService.YLWServiceModule.FileuploadGetSeq(security, ds);
        }

        public DataSet FileDownload(YlwSecurityJson security, DataSet ds)
        {
            return YLWService.YLWServiceModule.FileDownload(security, ds);
        }

        public string FileDownloadBase64(YlwSecurityJson security, DataSet ds)
        {
            return YLWService.YLWServiceModule.FileDownloadBase64(security, ds);
        }

        public string FileDelete(YlwSecurityJson security, DataSet ds)
        {
            return YLWService.YLWServiceModule.FileDelete(security, ds);
        }

        // /OpenPostDocx/{streamdata}의 형식으로 접속되면 호출되어 처리한다.
        public ReportData GetReportPost(string streamdata)
        {
            string rptPath = "";
            string rptName = "";

            try
            {
                string value = streamdata;

                JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
                ReportParam para = JsonConvert.DeserializeObject<ReportParam>(value, settings);

                RptAdjSLRptMain rptMain = new RptAdjSLRptMain();
                //StartupPath 경로주의!!
                rptMain.myPath = HttpContext.Current.Server.MapPath("~/bin");
                YLWService.Response rsp = rptMain.RptMain(para, ref rptPath, ref rptName);
                rptName = rptName.Replace("*", "_");  //파일이름에 * 있으면 파일복사시 에러남
                ReportData response = new ReportData() { Response = rsp };
                if (rsp.Result == 1)
                {
                    byte[] rptbyte = (byte[])MetroSoft.HIS.cFile.ReadBinaryFile(rptPath);
                    string rptText = Convert.ToBase64String(rptbyte);
                    response.ReportName = rptName;
                    response.ReportText = rptText;
                }
                return response;
            }
            catch (Exception ex)
            {
                YLWService.Response rsp = new YLWService.Response() { Result = -999, Message = ex.Message };
                ReportData response = new ReportData() { Response = rsp };
                return response;
            }
            finally
            {
                //다운로드후에 파일삭제
                Utils.DeleteFile(rptPath);
            }
        }

        public ReportData GetSaveReportPost(string streamdata)
        {
            string rptPath = "";
            string rptName = "";

            try
            {
                string value = streamdata;

                JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
                ReportParam para = JsonConvert.DeserializeObject<ReportParam>(value, settings);

                RptAdjSLXSaveReport rptMain = new RptAdjSLXSaveReport();
                //StartupPath 경로주의!!
                rptMain.myPath = HttpContext.Current.Server.MapPath("~/bin");
                YLWService.Response rsp = rptMain.RptMain(para, ref rptPath, ref rptName);
                ReportData response = new ReportData() { Response = rsp };
                if (rsp.Result == 1)
                {
                    byte[] rptbyte = (byte[])MetroSoft.HIS.cFile.ReadBinaryFile(rptPath);
                    string rptText = Convert.ToBase64String(rptbyte);
                    response.ReportName = rptName;
                    response.ReportText = rptText;
                }
                return response;
            }
            catch (Exception ex)
            {
                YLWService.Response rsp = new YLWService.Response() { Result = -999, Message = ex.Message };
                ReportData response = new ReportData() { Response = rsp };
                return response;
            }
            finally
            {
                //다운로드후에 파일삭제
                Utils.DeleteFile(rptPath);
            }
        }

        public DataSet SaveReportHistory(string streamdata)
        {
            string rptPath = "";
            string rptName = "";

            try
            {
                string value = streamdata;

                JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
                ReportParam para = JsonConvert.DeserializeObject<ReportParam>(value, settings);

                RptAdjSLRptMain rptMain = new RptAdjSLRptMain();
                //StartupPath 경로주의!!
                rptMain.myPath = HttpContext.Current.Server.MapPath("~/bin");
                YLWService.Response rsp = rptMain.RptMain(para, ref rptPath, ref rptName);
                rptName = rptName.Replace("*", "_");  //파일이름에 * 있으면 파일복사시 에러남
                ReportData response = new ReportData() { Response = rsp };
                if (rsp.Result != 1)
                {
                    throw new Exception(rsp.Message);
                }

                string fileSeq = "0";  //새파일생성
                YLWService.YlwSecurityJson sec1 = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
                sec1.companySeq = para.CompanySeq;
                // File Info
                FileInfo finfo = new FileInfo(rptPath);
                byte[] rptbyte = (byte[])MetroSoft.HIS.cFile.ReadBinaryFile(rptPath);
                string fileBase64 = Convert.ToBase64String(rptbyte);
                // File Info
                //fileSeq = YLWService.MTRServiceModule.CallMTRFileuploadGetSeq(sec1, finfo, fileBase64, "47820005", fileSeq);  // 이부분에서 오류남. CallMTRFileuploadGetSeq -> FileuploadGetSeq
                //fileSeq = YLWService.YLWServiceModule.FileuploadGetSeq(sec1, finfo, fileBase64, "47820005", fileSeq);
                fileSeq = YLWService.MTRServiceModule.CallMTRFileuploadGetSeq(sec1, finfo, fileBase64, "47820005", fileSeq);  // WebYlwPlugin_MetroSoft -> 일반 POST API 로 변경
                if (fileSeq == "")
                {
                    throw new Exception("보고서 업로드 실패");
                }

                YLWService.YlwSecurityJson security = YLWService.MTRServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersCS";
                security.methodId = "ReportHistorySave";
                security.companySeq = para.CompanySeq;
                security.certId = security.certId + "_1";  // securityType = 1 --> ylwhnpsoftgw_1
                security.securityType = 1;
                security.userId = para.UserID;

                DataSet ds = new DataSet();
                DataTable dt10 = ds.Tables.Add("DataBlock10");
                dt10.Columns.Add("AcptMgmtSeq");
                dt10.Columns.Add("ReSurvAsgnNo");
                dt10.Columns.Add("ReportType");
                dt10.Columns.Add("Seq");
                dt10.Columns.Add("FileName");
                dt10.Columns.Add("FileSeq");

                dt10.Clear();
                DataRow dr1 = dt10.Rows.Add();
                dr1["AcptMgmtSeq"] = para.AcptMgmtSeq;
                dr1["ReSurvAsgnNo"] = para.ReSurvAsgnNo;
                dr1["ReportType"] = para.ReportType;
                dr1["FileName"] = rptName;
                dr1["FileSeq"] = fileSeq;

                return YLWService.MTRServiceModule.CallMTRServiceCallPost(security, ds);
            }
            catch (Exception ex)
            {
                DataSet dsr = new DataSet();
                DataTable dtr = dsr.Tables.Add("ErrorMessage");
                dtr.Columns.Add("Status");
                dtr.Columns.Add("Message");
                DataRow dr = dtr.Rows.Add();
                dr["Status"] = "ERR";
                dr["Message"] = ex.Message;
                return dsr;
            }
        }

        public ReportData GetSaveRptHistoryPost(string streamdata)
        {
            string rptPath = "";
            string rptName = "";

            try
            {
                string value = streamdata;

                JsonSerializerSettings settings = new JsonSerializerSettings() { StringEscapeHandling = StringEscapeHandling.EscapeHtml };
                ReportParam para = JsonConvert.DeserializeObject<ReportParam>(value, settings);

                RptAdjSLXSaveReport rptMain = new RptAdjSLXSaveReport();
                //StartupPath 경로주의!!
                rptMain.myPath = HttpContext.Current.Server.MapPath("~/bin");
                YLWService.Response rsp = rptMain.RptHistoryPost(para, ref rptPath, ref rptName);
                ReportData response = new ReportData() { Response = rsp };
                if (rsp.Result == 1)
                {
                    byte[] rptbyte = (byte[])MetroSoft.HIS.cFile.ReadBinaryFile(rptPath);
                    string rptText = Convert.ToBase64String(rptbyte);
                    response.ReportName = rptName;
                    response.ReportText = rptText;
                }
                return response;
            }
            catch (Exception ex)
            {
                YLWService.Response rsp = new YLWService.Response() { Result = -999, Message = ex.Message };
                ReportData response = new ReportData() { Response = rsp };
                return response;
            }
            finally
            {
                //다운로드후에 파일삭제
                Utils.DeleteFile(rptPath);
            }
        }
    }

    internal class data_json
    {
        public string AssemblyName { get; set; }
        public string ClassName { get; set; }
        public string MethodName { get; set; }
        public object ParamData { get; set; }
        public object JSonData { get; set; }
    }
}

