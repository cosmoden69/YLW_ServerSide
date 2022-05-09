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
    public class RptAdjSLRptSurvRptPersDBLife
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersDBLife(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersDBLife";
                security.methodId = "Query";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock1");

                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");
                dt.Columns.Add("ReportType");

                dt.Clear();
                DataRow dr = dt.Rows.Add();

                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;
                dr["ReportType"] = para.ReportType;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "Start");

                string sSampleXSD = myPath + @"\보고서\출력설계_1536_서식_DB생명 종결보고서.xsd";
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
                DataTable dtBCnt = pds.Tables["DataBlock10"];
                DataTable dtBPrg = pds.Tables["DataBlock3"];
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
                dtBPrg.TableName = "DataBlock3";
                pds.Tables.Remove("DataBlock3");
                pds.Tables.Add(dtBPrg);

                string sSampleDocx = myPath + @"\보고서\출력설계_1536_서식_DB생명 종결보고서.docx";
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
                
                rptName = "종결보고서_인보험_DB생명(" + sfilename + ").docx";
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
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B3CureFrDt@");
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B5OthINsurCoNm@");
                    Table oTbl세부항목 = rUtil.GetTable(lstTable, "@B8ShrtCnts1@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B9PrgMgtDt@");

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (oTbl일자별확인사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl일자별확인사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl타사가입사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl타사가입사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (oTbl세부항목 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl세부항목, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
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
                    Table oTbl일자별확인사항 = rUtil.GetTable(lstTable, "@B3CureFrDt@");
                    Table oTbl타사가입사항 = rUtil.GetTable(lstTable, "@B5OthINsurCoNm@");
                    Table oTbl세부항목 = rUtil.GetTable(lstTable, "@B8ShrtCnts1@");
                    Table oTbl사고조사처리과정 = rUtil.GetTable(lstTable, "@B9PrgMgtDt@");


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
                            if (col.ColumnName == "IsrdTel") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvReqDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CclsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CmplExptFg")
                            {
                                if (sValue == "1") sValue = "☑ Y / □ N";
                                else if (sValue == "2") sValue = "□ Y / ☑ N";
                                else sValue = "□ Y / □ N";
                            }
                            if (col.ColumnName == "SurvAsgnTeamLeadOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpHP") sValue = Utils.TelNumber(sValue);
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
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
                                    if (col.ColumnName == "CureFrDt") sValue = (sValue.Trim() == "" ? "-" : Utils.DateFormat(sValue, "yyyy.MM.dd"));
                                    if (col.ColumnName == "CureToDt") sValue = (sValue.Trim() == "" ? "-" : Utils.DateFormat(sValue, "yyyy.MM.dd"));
                                    rUtil.ReplaceTableRow(oTbl일자별확인사항.GetRow(i + 1), sKey, sValue);
                                }
                                if (dtB.Columns.Contains("Gubun") && dr["Gubun"] + "" == "계약일")
                                {
                                    rUtil.TableRowBackcolor(oTbl일자별확인사항.GetRow(i + 1), "ABCDEF");
                                }
                            }
                        }
                    }

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
                            if (col.ColumnName == "S101_ShrtCnts1") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S102_ShrtCnts3") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S133_ShrtCnts1") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "S133_ShrtCnts2") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
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
                                    if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl타사가입사항.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        Table oTable = rUtil.GetTable(lstTable, "@B6RprtHed1@");
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add("고지 직접 이행후 자필서명", "", "1");
                            if (dtB.Rows.Count < 2) dtB.Rows.Add("약관 전달 여부", "", "2");
                            if (dtB.Rows.Count < 3) dtB.Rows.Add("면책 약관 주요내용 설명", "", "3");
                            if (dtB.Rows.Count < 4) dtB.Rows.Add("면책 판단 구비서류", "", "4");
                            if (dtB.Rows.Count < 5) dtB.Rows.Add("처리 적정성 여부", "", "5");
                            if (dtB.Rows.Count < 6) dtB.Rows.Add("상반되는 판례", "", "6");
                            if (dtB.Rows.Count < 7) dtB.Rows.Add("작성자 불이익 원칙 적용", "", "7");
                            if (dtB.Rows.Count < 8) dtB.Rows.Add("재검토 가능 요소", "", "8");

                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int pos = Utils.ToInt(dr["RprtNo"]);
                                if (pos < 1) continue;
                                TableRow trow = oTable.GetRow(pos);
                                sKey = "@B6RprtHed" + pos + "@";
                                sValue = dr["RprtHed"] + "";
                                rUtil.SetText(trow.GetCell(0), sKey, sValue);
                                sValue = dr["RprtRevwRslt"] + "";
                                if (sValue == "1") sValue = "☑ " + rUtil.GetText(trow.GetCell(1));
                                else sValue = "□ " + rUtil.GetText(trow.GetCell(1));
                                rUtil.SetPlaneText(trow.GetCell(1), sValue);
                                sValue = dr["RprtRevwRslt"] + "";
                                if (sValue == "2") sValue = "☑ " + rUtil.GetText(trow.GetCell(2));
                                else sValue = "□ " + rUtil.GetText(trow.GetCell(2));
                                rUtil.SetPlaneText(trow.GetCell(2), sValue);
                            }
                        }
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
                            if (col.ColumnName == "InvcAmtCof") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcAdjFee") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcDocuAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcCsltReqAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcTrspExps") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcIctvAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (oTbl세부항목 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "Amt1") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "ShrtCnts3") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTbl세부항목.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
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
