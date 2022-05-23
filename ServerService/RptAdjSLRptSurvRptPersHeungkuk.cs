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
    public class RptAdjSLRptSurvRptPersHeungkuk
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersHeungkuk(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersHeungkuk";
                security.methodId = "Query";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1535_서식_흥국화재 종결보고서.xsd";
                //string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                DataTable dtB2 = pds.Tables["DataBlock2"];
                DataTable dtBCnt = pds.Tables["DataBlock4"];
                DataTable dtBPrg = pds.Tables["DataBlock9"];
                for (int i = 0; i < dtBCnt.Rows.Count; i++)
                {
                    DataRow drow = dtBPrg.Rows.Add();
                    drow["Gubun"] = "계약일";
                    drow["CureFrDt"] = dtBCnt.Rows[i]["CtrtDt"];
                    drow["CureCnts"] = dtBCnt.Rows[i]["InsurPrdt"];
                    drow["VstHosp"] = dtB2.Rows[0]["InsurCo"];
                    drow["CureSeq"] = 0;
                }
                dtBPrg = dtBPrg.Select("", "CureFrDt, CureSeq").CopyToDataTable<DataRow>();
                dtBPrg.TableName = "DataBlock9";
                pds.Tables.Remove("DataBlock9");
                pds.Tables.Add(dtBPrg);

                string sSampleDocx = myPath + @"\보고서\출력설계_1535_서식_흥국화재 종결보고서.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                
                //DOCX 파일합치기 
                WordprocessingDocument wdoc = WordprocessingDocument.Open(sSample1Relt, true);
                MainDocumentPart mainPart = wdoc.MainDocumentPart;
                for (int ii = 0; ii < addFiles.Count; ii++)
                {
                    string addFile = addFiles[ii];
                    RptUtils.AppendFile(mainPart, addFile, (ii > 0 ? true : false));
                    Utils.DeleteFile(addFile);
                }
                mainPart.Document.Save();
                wdoc.Close();

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
                
                rptName = "종결보고서_인보험_흥국화재(" + sfilename + ").docx";
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
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B7OthINsurCoNm@");
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B9CureFrDt@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        Table oTable = rUtil.GetTable(lstTable, "@B4InsurPrdt@");
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            DataRow dr = dtB.Rows[0];
                            //drs = pds.Tables["DataBlock5"]?.Select("InsurNo = '" + dr["InsurNo"] + "'");
                            ////테이블의 끝에 추가
                            //rUtil.TableInsertRows(oTable, 7, 2, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl타사가입사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl타사가입사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl일자별확인사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl일자별확인사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl사고조사처리과정 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl사고조사처리과정, 1, dtB.Rows.Count - 1);
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
                    Table oTbl표지 = rUtil.GetTable(lstTable, "@B2LeadAdjuster@");
                    Table oTbl민원지수 = rUtil.GetTable(lstTable, "@B3CmplPnt1A@");
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B7OthINsurCoNm@");
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B9CureFrDt@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");

                    var db2SurvAsgnEmpManRegNo = ""; //조사자 손해사정등록번호
                    var db2SurvAsgnEmpAssRegNo = ""; //조사자 보조인 등록번호
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "SurvAsgnTeamLeadOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpHP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "GivObjRels")
                            {
                                if (sValue != "")
                                {
                                    sValue = "(" + sValue + ")";
                                }
                            }
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
                            if (col.ColumnName == "LeadAdjManRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "ChrgAdjAssRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "SurvAsgnEmpManRegNo")
                            {
                                if (sValue != "") db2SurvAsgnEmpManRegNo = sValue;
                            }
                            if (col.ColumnName == "SurvAsgnEmpAssRegNo")
                            {
                                if (sValue != "") db2SurvAsgnEmpAssRegNo = sValue;
                            }
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    if (db2SurvAsgnEmpManRegNo == "")
                    {
                        if (db2SurvAsgnEmpAssRegNo != "")
                        {
                            db2SurvAsgnEmpAssRegNo = "보조인 등록번호 : 제" + db2SurvAsgnEmpAssRegNo + "호";
                        }
                        rUtil.ReplaceTable(oTbl표지, "@db2SurvAsgnEmpRegNo@", db2SurvAsgnEmpAssRegNo);
                    }
                    else
                    {
                        db2SurvAsgnEmpManRegNo = "손해사정등록번호 : 제" + db2SurvAsgnEmpManRegNo + "호";
                        rUtil.ReplaceTable(oTbl표지, "@db2SurvAsgnEmpRegNo@", db2SurvAsgnEmpManRegNo);
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "CmplPnt1")
                            {
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1A@", "②");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1B@", "①");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1C@", "ⓞ");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1A@", "2");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1B@", "1");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1C@", "0");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt2")
                            {
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2A@", "②");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2B@", "①");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2C@", "ⓞ");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2A@", "2");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2B@", "1");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2C@", "0");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt3")
                            {
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3A@", "②");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3B@", "①");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3C@", "ⓞ");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3A@", "2");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3B@", "1");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3C@", "0");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt4")
                            {
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4A@", "②");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4B@", "①");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4C@", "ⓞ");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4A@", "2");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4B@", "1");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4C@", "0");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt5")
                            {
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5A@", "②");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5B@", "①");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5C@", "ⓞ");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5A@", "2");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5B@", "1");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5C@", "0");
                                continue;
                            }
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        Table oTable = rUtil.GetTable(lstTable, "@B4InsurPrdt@");
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTable(oTable, sKey, sValue);
                                }
                                string cnts = "";
                                string amts = "";
                                drs = pds.Tables["DataBlock5"]?.Select("InsurNo = '" + dr["InsurNo"] + "'");
                                if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock5"].Rows.Add() };
                                for (int j = 0; j < drs.Length; j++)
                                {
                                    DataRow dr1 = drs[j];
                                    if (j > 0)
                                    {
                                        cnts += " / "; amts += " / ";
                                    }
                                    cnts += dr1["CltrCnts"];
                                    amts += Utils.AddComma(dr1["InsurRegsAmt"]) + "원";
                                    //int rnum = j * 2 + 7;
                                    //rUtil.ReplaceTableRow(oTable.GetRow(rnum + 0), "@B5CltrCnts@", dr1["CltrCnts"] + "");
                                    //rUtil.ReplaceTableRow(oTable.GetRow(rnum + 1), "@B5InsurRegsAmt@", Utils.AddComma(dr1["InsurRegsAmt"]) + "원");
                                }
                                rUtil.ReplaceTable(oTable, "@B5CltrCnts@", cnts);
                                rUtil.ReplaceTable(oTable, "@B5InsurRegsAmt@", amts);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "S101_ShrtCnts1") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S102_ShrtCnts3") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S133_ShrtCnts1") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S133_ShrtCnts2") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "OthInsurCoNm");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl일자별확인사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "CureFrDt")
                                    {
                                        sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                        if (Utils.ConvertToString(dr["CureToDt"] + "").Replace(" ", "") != "") sValue += "\r\n~ " + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
                                    }
                                    rUtil.ReplaceTableRow(oTbl일자별확인사항.GetRow(i + 1), sKey, sValue);
                                }
                                if (dtB.Columns.Contains("Gubun") && dr["Gubun"] + "" == "계약일")
                                {
                                    rUtil.TableRowBackcolor(oTbl일자별확인사항.GetRow(i + 1), "ABCDEF");
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl사고조사처리과정 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTbl사고조사처리과정.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
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
