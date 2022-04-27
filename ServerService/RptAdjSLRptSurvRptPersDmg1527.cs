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
    public class RptAdjSLRptSurvRptPersDmg1527
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersDmg1527(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1527_서식_손해사정서 교부 동의 및 확인서_삼성화재.xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_1527_서식_손해사정서 교부 동의 및 확인서_삼성화재.docx";
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

                    Table oTblA = rUtil.GetTable(lstTable, "@db2IsrtOthTel@"); //요청사항 표

                    string IsrdGuidRcvMthd = "";
                    string IsrtGuidRcvMthd = "";
                    string BnfcGuidRcvMthd = "";
                    string db2IsrdOthTel = "";
                    string db2IsrtOthTel = "";
                    string db2BnfcOthTel = "";

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
                            //수신방법 - 수신방법
                            if (col.ColumnName == "IsrtGuidRcvMthd") { IsrtGuidRcvMthd = sValue; } //계약자
                            if (col.ColumnName == "IsrdGuidRcvMthd") { IsrdGuidRcvMthd = sValue; } //피보험자
                            if (col.ColumnName == "BnfcGuidRcvMthd") { BnfcGuidRcvMthd = sValue; } //수익자
                            //수신방법 - 받을곳  (ADJ_손해사정서수신방법)
                            if (col.ColumnName == "IsrtOthTel")
                            { //계약자
                                if (IsrtGuidRcvMthd == "300112001") { db2IsrtOthTel = "☑ 전화 (" + sValue + ")" + "\n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                                else if (IsrtGuidRcvMthd == "300112002") { db2IsrtOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n☑ 주소 (" + sValue + ")"; }
                                else if (IsrtGuidRcvMthd == "300112003") { db2IsrtOthTel = "□ 전화 \n☑ E-Mail (" + sValue + ")" + "\n□ 팩스 \n□ 주소 "; }
                                else if (IsrtGuidRcvMthd == "300112004") { db2IsrtOthTel = "□ 전화 \n□ E-Mail \n☑ 팩스 (" + sValue + ")" + "\n□ 주소 "; }
                                else { db2IsrtOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                            }
                            if (col.ColumnName == "IsrdOthTel")
                            { //피보험자
                                if (IsrdGuidRcvMthd == "300112001") { db2IsrdOthTel = "☑ 전화 (" + sValue + ")" + "\n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                                else if (IsrdGuidRcvMthd == "300112002") { db2IsrdOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n☑ 주소 (" + sValue + ")"; }
                                else if (IsrdGuidRcvMthd == "300112003") { db2IsrdOthTel = "□ 전화 \n☑ E-Mail (" + sValue + ")" + "\n□ 팩스 \n□ 주소 "; }
                                else if (IsrdGuidRcvMthd == "300112004") { db2IsrdOthTel = "□ 전화 \n□ E-Mail \n☑ 팩스 (" + sValue + ")" + "\n□ 주소 "; }
                                else { db2IsrdOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                            }
                            if (col.ColumnName == "BnfcOthTel")
                            { //수익자
                                if (BnfcGuidRcvMthd == "300112001") { db2BnfcOthTel = "☑ 전화 (" + sValue + ")" + "\n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                                else if (BnfcGuidRcvMthd == "300112002") { db2BnfcOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n☑ 주소 (" + sValue + ")"; }
                                else if (BnfcGuidRcvMthd == "300112003") { db2BnfcOthTel = "□ 전화 \n☑ E-Mail (" + sValue + ")" + "\n□ 팩스 \n□ 주소 "; }
                                else if (BnfcGuidRcvMthd == "300112004") { db2BnfcOthTel = "□ 전화 \n□ E-Mail \n☑ 팩스 (" + sValue + ")" + "\n□ 주소 "; }
                                else { db2BnfcOthTel = "□ 전화 \n□ E-Mail \n□ 팩스 \n□ 주소 "; }
                            }
                            //수신방법 - 개인정보 제공 동의  (ADJ_동의여부)
                            if (col.ColumnName == "IsrtSstvInfoAgrFg")
                            { //계약자
                                if (sValue == "300113001") { sValue = "☑ 동의함 / □ 부동의"; }
                                else { sValue = "□ 동의함 / ☑ 부동의"; }
                            }
                            if (col.ColumnName == "IsrdSstvInfoAgrFg")
                            { //피보험자
                                if (sValue == "300113001") { sValue = "☑ 동의함 / □ 부동의"; }
                                else { sValue = "□ 동의함 / ☑ 부동의"; }
                            }
                            if (col.ColumnName == "BnfcSstvInfoAgrFg")
                            { //수익자
                                if (sValue == "300113001") { sValue = "☑ 동의함 / □ 부동의"; }
                                else { sValue = "□ 동의함 / ☑ 부동의"; }
                            }

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    rUtil.ReplaceTable(oTblA, "@db2IsrtOthTel@", db2IsrtOthTel); //수신방법 (계약자)
                    rUtil.ReplaceTable(oTblA, "@db2IsrdOthTel@", db2IsrdOthTel); //수신방법 (피보험자)
                    rUtil.ReplaceTable(oTblA, "@db2BnfcOthTel@", db2BnfcOthTel); //수신방법 (수익자)

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
