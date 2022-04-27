using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.Schema;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

using YLWService;
using YLWService.Extensions;

namespace YLW_WebService.ServerSide
{
    public class RptAdjSLRptSurvRptPersDmg1525
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersDmg1525(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisDmgRptAgrPrint";
                security.methodId = "Print";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock1");

                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");

                dt.Clear();
                DataRow dr = dt.Rows.Add();

                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "Start");

                string sSampleXSD = myPath + @"\보고서\출력설계_1525_서식_손해사정서 교부 동의 및 확인서_공통.xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_1525_서식_손해사정서 교부 동의 및 확인서_공통.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "손해사정서교부확인서_인보험(" + sfilename + ").docx";
                rptPath = sSample1Relt;
                //System.Diagnostics.Process process = System.Diagnostics.Process.Start(sSample1Relt);
                //Utils.BringWindowToTop(process.Handle);

                return new Response() { Result = 1, Message = "OK" };
            }
            catch (Exception ex)
            {
                return new Response() { Result = -99, Message = ex.Message };
            }
        }

        private string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
        {
            string sRet = "";

            if (!File.Exists(sDocFile)) return RptUtils.GetMessage("원본파일(word)이 존재하지 않습니다.", sDocFile);
            if (!File.Exists(sXSDFile)) return RptUtils.GetMessage("XSD파일이 존재하지 않습니다.", sXSDFile);

            DataTable dtB = null;
            DataRow[] drs = null;
            string sPrefix = "";
            string sKey = "";
            string sValue = "";

            try
            {
                System.IO.File.Copy(sDocFile, sWriteFile, true);

                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(sWriteFile, true))
                {
                    MainDocumentPart mDoc = wDoc.MainDocumentPart;
                    Document doc = mDoc.Document;
                    RptUtils rUtil = new RptUtils(mDoc);

                    IEnumerable<Table> lstTable = doc.Body.Elements<Table>();

                    //테이블에 행을 추가하고 일단 저장
                    // Save
                    doc.Save();
                    wDoc.Close();
                }

                //=== repalce ===
                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(sWriteFile, true))
                {
                    MainDocumentPart mDoc = wDoc.MainDocumentPart;
                    Document doc = mDoc.Document;
                    RptUtils rUtil = new RptUtils(mDoc);

                    List<Table> lstTable = doc.Body.Elements<Table>()?.ToList();

                    Table oTblA = rUtil.GetTable(lstTable, "@db2OthTel@"); //손해사정서 교부용 발송 정보 받을 곳 표
                    Table oTblB = rUtil.GetTable(lstTable, "@db2SstvInfoAgrFg@"); //민감정보 및 고유식별정보 제공 표

                    string db2OthTel = ""; //손해사정서 교부용 발송 정보 받을 곳 
                    string IsrtSstvInfoAgrFg = ""; //계약자 민감정보 및 고유식별정보 제공 여부
                    string DmdRgtSstvInfoAgrFg = ""; //청구권자 민감정보 및 고유식별정보 제공 여부
                    string db2SstvInfoAgrFg = ""; //계약자 및 청구권자 민감정보 및 고유식별정보 제공 여부

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "IsrdTel") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "IsrtTel") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "BnfcTel") sValue = Utils.TelNumber(sValue);
                            //손해사정서 교부여부 (ADJ_동의여부)
                            if (col.ColumnName == "LosrptDlvrFg_Isrt")
                            { //계약자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            if (col.ColumnName == "LosrptDlvrFg_Isrd")
                            { //피보험자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            if (col.ColumnName == "LosrptDlvrFg_DmdRgt")
                            { //청구권자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            //손해사정서 교부용 발송 정보 (ADJ_손해사정서수신방법)
                            if (col.ColumnName == "IsrtGuidRcvMthd")
                            { //계약자
                                if (sValue == "300112001") { sValue = "☑ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112002") { sValue = "□ 문자메세지 ☑ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112003") { sValue = "□ 문자메세지 □ 우편 ☑ 전자우편 □ 팩스"; }
                                else if (sValue == "300112004") { sValue = "□ 문자메세지 □ 우편 □ 전자우편 ☑ 팩스"; }
                                else { sValue = "□ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                            }
                            if (col.ColumnName == "IsrdGuidRcvMthd")
                            { //피보험자
                                if (sValue == "300112001") { sValue = "☑ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112002") { sValue = "□ 문자메세지 ☑ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112003") { sValue = "□ 문자메세지 □ 우편 ☑ 전자우편 □ 팩스"; }
                                else if (sValue == "300112004") { sValue = "□ 문자메세지 □ 우편 □ 전자우편 ☑ 팩스"; }
                                else { sValue = "□ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                            }
                            if (col.ColumnName == "DmdGuidRcvMthd")
                            { //청구권자
                                if (sValue == "300112001") { sValue = "☑ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112002") { sValue = "□ 문자메세지 ☑ 우편 □ 전자우편 □ 팩스"; }
                                else if (sValue == "300112003") { sValue = "□ 문자메세지 □ 우편 ☑ 전자우편 □ 팩스"; }
                                else if (sValue == "300112004") { sValue = "□ 문자메세지 □ 우편 □ 전자우편 ☑ 팩스"; }
                                else { sValue = "□ 문자메세지 □ 우편 □ 전자우편 □ 팩스"; }
                            }
                            //개인정보 수집이용 동의(ADJ_동의여부)
                            if (col.ColumnName == "IsrtPrsInfoAgrFg")
                            { //계약자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            if (col.ColumnName == "IsrdPrsInfoAgrFg")
                            { //피보험자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            if (col.ColumnName == "DmdRgtPrsInfoAgrFg")
                            { //청구권자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            //손해사정서 교부용 발송 정보 받을 곳
                            if (col.ColumnName == "IsrtOthTel")
                            { //계약자
                                if (sValue != "") { db2OthTel = db2OthTel + "계약자(" + sValue + ")"; }
                            }
                            if (col.ColumnName == "IsrdOthTel")
                            { //피보험자
                                if (sValue != "") { db2OthTel = db2OthTel + ", 피보험자(" + sValue + ")"; }
                            }
                            if (col.ColumnName == "DmdRgtOthTel")
                            { //청구권자
                                if (sValue != "") { db2OthTel = db2OthTel + ", 청구권자(" + sValue + ")"; }
                            }

                            //민감정보 및 고유식별정보 제공
                            if (col.ColumnName == "IsrdSstvInfoAgrFg")
                            { //피보험자
                                if (sValue == "300113001") { sValue = "☑ 동의함 □ 동의안함"; }
                                else { sValue = "□ 동의함 ☑ 동의안함"; }
                            }
                            if (col.ColumnName == "IsrtSstvInfoAgrFg")
                            { //계약자
                                IsrtSstvInfoAgrFg = sValue;
                            }
                            if (col.ColumnName == "DmdRgtSstvInfoAgrFg")
                            { //청구권자
                                DmdRgtSstvInfoAgrFg = sValue;
                            }

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    //손해사정서 교부용 발송 정보 받을 곳
                    rUtil.ReplaceTable(oTblA, "@db2OthTel@", db2OthTel);

                    //계약자 및 청구권자 민감정보 및 고유식별정보 제공 여부
                    if ((IsrtSstvInfoAgrFg == "300113001") && (DmdRgtSstvInfoAgrFg == "300113001"))
                    {
                        db2SstvInfoAgrFg = "☑ 동의함 □ 동의안함";
                    }
                    else
                    {
                        db2SstvInfoAgrFg = "□ 동의함 ☑ 동의안함";
                    }
                    rUtil.ReplaceTable(oTblB, "@db2SstvInfoAgrFg@", db2SstvInfoAgrFg);

                    doc.Save();
                    wDoc.Close();
                }
            }
            catch (Exception ec)
            {
                sRet = RptUtils.GetMessage(ec.Message, ec.ToString());
            }

            return sRet;
        }
    }
}
