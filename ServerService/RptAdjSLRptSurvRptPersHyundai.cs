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
    public class RptAdjSLRptSurvRptPersHyundai
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPersHyundai(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtMngPersHyundai";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1534_서식_현대해상 종결보고서.xsd";
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
                DataTable dtBCnt = pds.Tables["DataBlock3"];
                DataTable dtBPrg = pds.Tables["DataBlock7"];
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
                dtBPrg.TableName = "DataBlock7";
                pds.Tables.Remove("DataBlock7");
                pds.Tables.Add(dtBPrg);

                string sSampleDocx = myPath + @"\보고서\출력설계_1534_서식_현대해상 종결보고서.docx";
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
                rptName = "종결보고서_인보험_현대해상(" + sfilename + ").docx";
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
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B3InsurPrdt@");
                    Table oTbl손해액범위조사 = rUtil.GetTable(lstTable, "@B7CureFrDt@");
                    Table oTbl조사일정표요약 = rUtil.GetTable(lstTable, "@B8PrgMgtDt@");
                                         
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    int B3RowCnt = 0;
                    B3RowCnt = dtB.Rows.Count+1;
                    if (dtB != null)
                    {
                        if (oTbl보험계약사항 != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableInsertRow(oTbl보험계약사항, 1, 1);
                            }
                        }
                    }
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (oTbl보험계약사항 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl보험계약사항, 1+ B3RowCnt, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl손해액범위조사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl손해액범위조사, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (oTbl조사일정표요약 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl조사일정표요약, 1, dtB.Rows.Count - 1);
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
                    Table oTbl조사요약 = rUtil.GetTable(lstTable, "@B5S132_LongCnts1@");

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
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "IsrdRegno1") sValue = "(" + sValue + " - *)";

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
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "InsurPrdt");
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
                                    if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    int B3RowCnt = 0;
                    B3RowCnt = dtB.Rows.Count + 1;
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CltrCnts");
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
                                    if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "InsurDmndAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1 + B3RowCnt), sKey, sValue);
                                }
                            }
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
                            if (sKey == "@B5S131_Amt1@")
                            {
                                TableRow trow = oTbl조사요약.GetRow(2);
                                rUtil.SetPlaneText(trow.GetCell(5), sValue);
                                continue;
                            }
                            if (sKey == "@B5S132_Amt1@")
                            {
                                TableRow trow = oTbl조사요약.GetRow(3);
                                rUtil.SetPlaneText(trow.GetCell(5), sValue);
                                continue;
                            }
                            if (sKey == "@B5S133_Amt1@")
                            {
                                TableRow trow = oTbl조사요약.GetRow(4);
                                rUtil.SetPlaneText(trow.GetCell(5), sValue);
                                continue;
                            }
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
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

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CureFrDt");
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
                                    if (col.ColumnName == "Gubun")
                                    {
                                        if (sValue == "계약전") sValue = "4-1\r\n" + sValue;
                                        if (sValue == "계약일") sValue = "4-2\r\n" + sValue;
                                        if (sValue == "계약후") sValue = "4-3\r\n" + sValue;
                                    }
                                    if (col.ColumnName == "CureFrDt")
                                    {
                                        sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                        if (Utils.ConvertToString(dr["CureToDt"] + "").Replace(" ", "") != "") sValue += "\r\n~ " + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
                                        if (Utils.ToInt(dr["OutHospDay"]) > 0) sValue += "\r\n(" + dr["OutHospDay"] + "회 통원)";
                                        if (Utils.ToInt(dr["InHospDay"]) > 0) sValue += "\r\n(" + dr["InHospDay"] + "일 입원)";
                                    }
                                    if (col.ColumnName == "BfGivCnts")
                                    {
                                        sValue = sValue + "";
                                        rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                                if (dtB.Columns.Contains("Gubun") && dr["Gubun"] + "" == "계약일")
                                {
                                    rUtil.TableRowBackcolor(oTable.GetRow(i + 1), "ABCDEF");
                                }
                            }
                            string txt1 = null;
                            int spos = 1;
                            int pos = spos;
                            for (int i = spos; i <= spos + dtB.Rows.Count - 1; i++)
                            {
                                sValue = rUtil.GetText(oTable.GetRow(i).GetCell(0));
                                if (txt1 != null && txt1 != sValue)
                                {
                                    rUtil.TableMergeCellsV(oTable, 0, pos, i - 1);
                                    pos = i;
                                }
                                txt1 = sValue;
                            }
                            if (dtB.Rows.Count > 0) rUtil.TableMergeCellsV(oTable, 0, pos, spos + dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "PrgMgtDt");
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
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
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
