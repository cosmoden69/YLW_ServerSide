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
    public class RptAdjSLRptSurvRptPersDmg
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersDmg(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisDmgRprtPersPrint";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1522_서식_손해사정서.xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_1522_서식_손해사정서.docx";
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
                rptName = "손해사정서_인보험(" + sfilename + ").docx";
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B2InsurPrdt@");
                    Table oTbl청구사항 = rUtil.GetTable(lstTable, "@B3DmndInsurNo@");

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (oTbl계약사항 != null)
                        {
                            rUtil.TableInsertRow(oTbl계약사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (oTbl청구사항 != null)
                        {
                            rUtil.TableInsertRow(oTbl청구사항, 1, dtB.Rows.Count - 1);
                        }
                    }

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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B2InsurPrdt@");
                    Table oTbl청구사항 = rUtil.GetTable(lstTable, "@B3DmndInsurNo@");


                    
                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    string IsrdPrsInfoAgrFg = "";
                    string InsurGivTypCd = "";
                    string InsurKeepCd = "";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        IsrdPrsInfoAgrFg = dr["IsrdPrsInfoAgrFg"] + "";  //피보험자 개인정보/민감정보 제공 여부
                        InsurGivTypCd = dr["InsurGivTypCd"] + "";        //보험지급유형
                        InsurKeepCd = dr["InsurKeepCd"] + "";            //보험유지구분
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "AcptDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "CclsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "EmpCellPhone") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjPhoto" || col.ColumnName == "ChrgAdjPhoto" || col.ColumnName == "SealPhotoLead" || col.ColumnName == "SealPhotoEmp")
                            {
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.ReplaceInternalImage(sKey, img);
                                }
                                catch { }
                                continue;
                            }
                            if (col.ColumnName == "LeadAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "ChrgAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "BistLicSerl")
                            {
                                if (sValue != "") sValue = "보 조 인 등록번호 : 제 " + sValue + " 호";
                            }
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "CtrtDt") sValue = Utils.DateConv(sValue, ".");
                                rUtil.ReplaceTableRow(oTbl계약사항.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "InHospDay")
                                {
                                    sValue = (dr["DmndInsurNo"] + "" == "" ? "" : (Utils.ToInt(sValue) > 0 ? "입원" : "통원"));
                                }
                                if (col.ColumnName == "CureFrDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "CureToDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "AcdtNo") sValue = Utils.DateConv(sValue, ".");
                                rUtil.ReplaceTableRow(oTbl청구사항.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (IsrdPrsInfoAgrFg == "300113001")  //개인정보제공동의 - 동의
                        {
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts1@", dr["SurvCnts_A310"] + "");       //청구내용
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts2@", dr["SurvCnts_A320"] + "");       //확인내용
                            rUtil.ReplaceTables(lstTable, "@db4IsrdPrsInfoAgrFg@", "");                      //동의여부
                        }
                        else if (IsrdPrsInfoAgrFg == "300113002")  //개인정보제공동의 - 부동의
                        {
                            sValue = "현장 확인결과, 해당 보험약관에 따라 약관상 해당여부 및 계약내용의 변경 검토가 필요함.";
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts1@", "");                             //청구내용
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts2@", "");                             //확인내용
                            rUtil.ReplaceTables(lstTable, "@db4IsrdPrsInfoAgrFg@", sValue);                  //동의여부
                        }
                        else
                        {
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts1@", "");                             //청구내용
                            rUtil.ReplaceTables(lstTable, "@db4SurvCnts2@", "");                             //확인내용
                            rUtil.ReplaceTables(lstTable, "@db4IsrdPrsInfoAgrFg@", "");                      //동의여부
                        }

                        rUtil.ReplaceTables(lstTable, "@db4SurvCnts3@", dr["SurvCnts_A410"] + "");           //보험금지급여부
                        rUtil.ReplaceTables(lstTable, "@db4SurvCnts4@", dr["SurvCnts_A420"] + "");           //계약전알릴의무위반여부
                        sValue = "";
                        if (InsurGivTypCd == "300048001" && InsurKeepCd == "300047001")  //지급 & 계약유지
                        {
                            sValue = "금번 청구 및 그 외 계약 전 알릴의무 위반사실 확인되지 않으며, 약관 상 지급 기준에 부합하는 것으로 판단되나, 라이나생명 최종 검토에 따라 추가적인 확인 및 결과가 변경될 수 있습니다.";
                        }
                        if (InsurGivTypCd == "300048001" && InsurKeepCd == "300047002")  //지급 & 계약해지
                        {
                            sValue = "금번 청구와 관련하여 계약 전 알릴의무 위반사실 확인되지 않으며, 약관 상 지급 기준에 부합하는 것으로 판단되나, 그 외 계약 전 알릴의무 위반사실 확인되는 바, 해당 계약애 대한 검토가 필요할 것으로 판단되며, 라이나생명 최종 검토에 따라 추가적인 확인 및 결과가 변경될 수 있습니다.";
                        }
                        if (InsurGivTypCd == "300048002" && InsurKeepCd == "300047001")  //부지급 & 계약유지
                        {
                            sValue = "계약 전 알릴의무 위반사실 확인되지 않으나, 약관 상 지급 기준에 부합하지 않는 바, 청구 보험금 부지급 및 계약유지 검토가 필요할 것으로 판단되며, 라이나생명 최종 검토에 따라 추가적인 확인 및 결과가 변경될 수 있습니다.";
                        }
                        if (InsurGivTypCd == "300048002" && InsurKeepCd == "300047002")  //부지급 & 계약해지
                        {
                            sValue = "계약 전 알릴의무 위반사실 확인되는 바, 청구 보험금 부지급 검토 및 해당 계약에 대한 검토가 필요할 것으로 판단되며, 라이나생명 최종 검토에 따라 추가적인 확인 및 결과가 변경될 수 있습니다.";
                        }
                        rUtil.ReplaceTables(lstTable, "@db4GivTypCdKeepCd@", sValue);

                        rUtil.ReplaceTables(lstTable, "@db4SurvCnts5@", dr["SurvCnts_A500"] + "");           //약관규정 및 관련법규
                    }

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
