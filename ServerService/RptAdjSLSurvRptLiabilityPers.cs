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
    public class RptAdjSLSurvRptLiabilityPers
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityPers(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintPers";
                security.methodId = "QueryCclsPrtLiabilityPer";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2553_서식_종결보고서(배책-대인).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2553_서식_종결보고서(배책-대인).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSample1Docx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "종결보고서_배책_대인(" + sfilename + ").docx";
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

                    //Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B6PrgMgtDt@"); // 사고처리과정표
                    
                    Table oTableC = rUtil.GetTable(lstTable, "@B7ExpsLosReq92@");

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        sKey = "@B1AcdtCnts@";
                        Table oTblA = rUtil.GetTable(lstTable, sKey);
                        sKey = "@B2AcdtPictImage@";
                        TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                        Table oTableA = oTblARow?.GetCell(1).Elements<Table>().FirstOrDefault();
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableA, 0, 1);
                                rUtil.TableAddRow(oTableA, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CureSerl");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "ShdnSerl");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
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

                    //Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B6PrgMgtDt@"); // 사고처리과정표
                    Table oTableC = rUtil.GetTable(lstTable, "@B7ExpsLosReq92@");

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    sKey = "@B1AcdtCnts@";
                    Table oTblA = rUtil.GetTable(lstTable, sKey);
                    sKey = "@B2AcdtPictImage@";
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    Table oTableA = oTblARow?.GetCell(1).Elements<Table>().FirstOrDefault();

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "DeptName") sValue = (sValue == "" ? "-" : sValue);
                            if (col.ColumnName == "EmpWorkAddress") sValue = (sValue == "" ? "-" : sValue);
                            if (col.ColumnName == "DeptPhone") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "DeptFax") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CtrtDt") if (sValue != "") { sValue = Utils.DateConv(sValue, "."); } //sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") if (sValue != "") { sValue = Utils.DateConv(sValue, "."); } //sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            //if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue); 삭제
                            //if (col.ColumnName == "DiSubTotReq") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DiTotReq") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DiSubTotAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DiTotAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DiTotAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DiTotDmg") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSelfBearAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DIGivInsurAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DiGivInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "VitmCureFrDt") if (sValue != "") { sValue = Utils.DateConv(sValue, "."); } //sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "VitmCureToDt") if (sValue != "") { sValue = Utils.DateConv(sValue, "."); } //sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtCausCatg2Nm") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "IsrtTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "IsrdTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "VitmTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "ExpsLosAmtTot1") if (sValue != "") {sValue = Utils.ConvertToString(Utils.AddComma(sValue)) + "￦"; }
                            if (col.ColumnName == "ExpsLosAmtTot2") if (sValue != "") { sValue = Utils.ConvertToString(Utils.AddComma(sValue)) + "￦"; }
                            if (col.ColumnName == "SealPhoto" || col.ColumnName == "ChrgAdjPhoto" || col.ColumnName == "LeadAdjPhoto")
                            {
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.ReplaceInternalImage(sKey, img);
                                }
                                catch { }
                                continue;
                            }
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (oTableA != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2;
                                int rmdr = i % 2;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableA.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableA.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }
                    
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (!dtB.Columns.Contains("VitmNglgRatePer")) dtB.Columns.Add("VitmNglgRatePer");
                        dr["VitmNglgRatePer"] = dr["VitmNglgRate"] + "%";
                        
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CureSerl");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                //DataRow dr = dtB.Rows[0];

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");//치료기간Fr
                                    if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");//치료기간To
                                    if (col.ColumnName == "DmndAmt") sValue = Utils.AddComma(sValue);//청구액
                                    if (col.ColumnName == "AdmtAmt") sValue = Utils.AddComma(sValue);//인정액
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                    

                                }
                            }
                        }
                    }


                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                //DataRow dr = dtB.Rows[0];

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "ShdnSerl");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                //DataRow dr = dtB.Rows[0];

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ShdnFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "ShdnToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "PayCost") sValue = Utils.AddComma(sValue);//노임단가
                                    if (col.ColumnName == "ShdnLosAmt") sValue = Utils.AddComma(sValue);//휴업손해액
                                    if (col.ColumnName == "AplcRate")
                                    {
                                        if ((sValue != null)&&(Utils.ToInt(sValue) > 0)) {
                                            sValue = sValue + "%";
                                        }
                                    }

                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);


                                }
                            }
                        }
                    }
                    

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        //1.치료비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        double dReq = 0;
                        double dAmt = 0;
                        string sExpsCmnt = "";
                        string sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        TableRow oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq1@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt1@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss1@", sExpsBss);
                        }

                        //2.휴업손해
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq2@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt2@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss2@", sExpsBss);
                        }

                        //3.상실수익
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq3@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt3@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss3@", sExpsBss);
                        }

                        //4.향후치료비
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        TableRow oRowBase = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq4@");
                        int rIdx1 = -1;
                        int rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableC, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsSubHed4@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq4@", Utils.AddComma(drs[i]["ExpsLosReq"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt4@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt4@", drs[i]["ExpsCmnt"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsBss4@", drs[i]["ExpsBss"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTableC, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTableC, oRow);
                        }
                        rUtil.TableMergeCellsV(oTableC, 0, rIdx1, rIdx2);

                        //5.개호비
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq5@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt5@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss5@", sExpsBss);
                        }

                        //6.기타손해
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        oRowBase = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq6@");
                        rIdx1 = -1;
                        rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableC, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsSubHed6@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq6@", Utils.AddComma(drs[i]["ExpsLosReq"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt6@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt6@", drs[i]["ExpsCmnt"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B7ExpsBss6@", drs[i]["ExpsBss"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTableC, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTableC, oRow);
                        }
                        rUtil.TableMergeCellsV(oTableC, 0, rIdx1, rIdx2);

                        //91.소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq91@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt91@", Utils.AddComma(dAmt));
                        }

                        //7.과실부담금
                        drs = dtB?.Select("ExpsGrp = 7");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq7@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq7@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt7@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt7@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss7@", sExpsBss);
                        }

                        //8.위자료
                        drs = dtB?.Select("ExpsGrp = 8");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq8@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq8@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt8@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt8@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss8@", sExpsBss);
                        }

                        //92.합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq92@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt92@", Utils.AddComma(dAmt));
                        }

                        //9.자기부담금
                        drs = dtB?.Select("ExpsGrp = 9");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq9@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq9@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt9@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsCmnt9@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsBss9@", sExpsBss);
                        }

                        //93.예상지급보험금
                        drs = dtB?.Select("ExpsGrp = 93");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B7ExpsLosReq93@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosReq93@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B7ExpsLosAmt93@", Utils.AddComma(dAmt));
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
