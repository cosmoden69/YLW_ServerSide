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
    public class RptAdjSLRptSurvRptPersMGLoss
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersMGLoss(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersMGLoss";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1539_서식_MG손해 종결보고서(일반).xsd";
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
                DataTable dtBCnt = pds.Tables["DataBlock5"];
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

                string sSampleDocx = myPath + @"\보고서\출력설계_1539_서식_MG손해 종결보고서(일반).docx";
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
                
                rptName = "종결보고서_인보험_MG손해(" + sfilename + ").docx";
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B5InsurPrdt@");
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B9CureFrDt@");
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B10OthINsurCoNm@");
                    Table oTbl첨부자료 = rUtil.GetTable(lstTable, "@B12FileSavSerl@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B14PrgMgtDt@");

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl계약사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            DataRow dr = dtB.Rows[0];
                            drs = pds.Tables["DataBlock6"]?.Select("InsurNo = '" + dr["InsurNo"] + "'");
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl계약사항, 5, drs.Length - 1);
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
                        if (oTbl타사가입사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl타사가입사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (oTbl첨부자료 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl첨부자료, 1, dtB.Rows.Count - 1);
                        }
                    }


                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
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
                    Table oTbl민원지수 = rUtil.GetTable(lstTable, "@B3CmplPnt1A@");
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B5InsurPrdt@");
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B9CureFrDt@");
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B10OthINsurCoNm@");
                    Table oTbl첨부자료 = rUtil.GetTable(lstTable, "@B12FileSavSerl@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B14PrgMgtDt@");

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "AcptDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "DlyRprtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");

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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    double db3CmplPntSum = 0;
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
                                db3CmplPntSum += Utils.ToDouble(sValue);
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1A@", "○");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1B@", "○");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1C@", "○");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1A@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1B@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt1C@", "");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt2")
                            {
                                db3CmplPntSum += Utils.ToDouble(sValue);
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2A@", "○");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2B@", "○");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2C@", "○");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2A@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2B@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt2C@", "");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt3")
                            {
                                db3CmplPntSum += Utils.ToDouble(sValue);
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3A@", "○");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3B@", "○");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3C@", "○");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3A@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3B@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt3C@", "");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt4")
                            {
                                db3CmplPntSum += Utils.ToDouble(sValue);
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4A@", "○");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4B@", "○");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4C@", "○");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4A@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4B@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt4C@", "");
                                continue;
                            }
                            if (col.ColumnName == "CmplPnt5")
                            {
                                db3CmplPntSum += Utils.ToDouble(sValue);
                                if (sValue == "3") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5A@", "○");
                                if (sValue == "2") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5B@", "○");
                                if (sValue == "1") rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5C@", "○");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5A@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5B@", "");
                                rUtil.ReplaceTable(oTbl민원지수, "@B3CmplPnt5C@", "");
                                continue;
                            }
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    rUtil.ReplaceTextAllParagraph(doc, "@db3CmplPntSum@", Utils.AddComma(db3CmplPntSum));

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "IsrdTel") sValue = Utils.TelNumber(sValue);
                            rUtil.ReplaceTable(oTbl계약사항, sKey, sValue);
                        }
                        drs = pds.Tables["DataBlock6"]?.Select("InsurNo = '" + dr["InsurNo"] + "'");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock6"].Rows.Add() };
                        for (int j = 0; j < drs.Length; j++)
                        {
                            DataRow dr1 = drs[j];
                            int rnum = j + 5;
                            rUtil.ReplaceTableRow(oTbl계약사항.GetRow(rnum), "@B6CltrCnts@", dr1["CltrCnts"] + "");
                            rUtil.ReplaceTableRow(oTbl계약사항.GetRow(rnum), "@B6InsurRegsAmt@", Utils.AddComma(dr1["InsurRegsAmt"]) + "원");
                        }
                        string txt1 = null;
                        int spos = 5;
                        int pos = spos;
                        for (int j = spos; j <= spos + drs.Length - 1; j++)
                        {
                            sValue = rUtil.GetText(oTbl계약사항.GetRow(j).GetCell(0));
                            if (txt1 != null && txt1 != sValue)
                            {
                                rUtil.TableMergeCellsV(oTbl계약사항, 0, pos, j - 1);
                                pos = j;
                            }
                            txt1 = sValue;
                        }
                        if (drs.Length > 0) rUtil.TableMergeCellsV(oTbl계약사항, 0, pos, spos + drs.Length - 1);
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
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
                                        if (Utils.ConvertToString(dr["CureToDt"] + "").Replace(" ", "") != "") sValue += "\r\n~\r\n" + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
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
                        if (oTbl타사가입사항 != null)
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
                                    rUtil.ReplaceTableRow(oTbl타사가입사항.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (oTbl첨부자료 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTbl첨부자료.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
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
