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
    public class RptAdjSLSurvMidRptLiabilityPersNH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvMidRptLiabilityPersNH_S(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintPersNH";
                security.methodId = "QueryCclsPrtLiabilityPerNH";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2575_서식_농협_진행보고서(배책-대인, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2575_서식_농협_진행보고서(배책-대인, 간편).docx";
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
                rptName = "진행보고서_배책_대인, 간편(" + sfilename + ").docx";
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

                    dtB = pds.Tables["DataBlock5"];
                    if (dtB != null)
                    {
                        sKey = "@B5FileNo@";
                        Table oTableA = rUtil.GetTable(lstTable, sKey);
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableA, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        sKey = "@B10PrgMgtDt@";
                        Table oTableB = rUtil.GetTable(lstTable, sKey);
                        if (oTableB != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableB, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //피해자
                    dtB = pds.Tables["DataBlock12"];
                    if (dtB != null)
                    {
                        sKey = "@B12VitmNm@";
                        Table oTableC = rUtil.GetTable(lstTable, sKey);
                        if (oTableC != null)
                        {
                            //테이블의 끝에 추가
                            dtB.Rows.Count.Equals(dtB.Columns);
                            rUtil.TableInsertRows(oTableC, 2, 2, dtB.Rows.Count - 1);
                            //rUtil.TableInsertRows(oTableC, (dtB.Rows.Count * 2) + 4, 3, dtB.Rows.Count - 1);
                        }
                    }

                    //기왕력검토
                    int B12RowsCount = dtB.Rows.Count;
                    dtB = pds.Tables["DataBlock4"];
                    if (dtB != null)
                    {
                        sKey = "@B4MedHstr@";
                        Table oTableD = rUtil.GetTable(lstTable, sKey);
                        if (oTableD != null)
                        {
                            //테이블의 끝에 추가
                            //rUtil.TableInsertRows(oTableC, 2, 2, dtB.Rows.Count - 1);
                            //rUtil.TableInsertRows(oTableC, (dtB.Rows.Count * 2) + 4, 3, dtB.Rows.Count - 1);
                            rUtil.TableInsertRows(oTableD, (B12RowsCount * 2) + 4, 3, dtB.Rows.Count - 1);
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

                    Table oTableH = rUtil.GetTable(lstTable, "@B13ExpsLosAmt92@");

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
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            //if (col.ColumnName == "ObjSelfBearAmt") sValue = Utils.AddComma(sValue);


                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            //4. 추정지급보험금 표
                            if (col.ColumnName == "DiSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DIGivInsurAmt") sValue = Utils.AddComma(sValue);
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
                            if (col.ColumnName == "VitmNglgRate") sValue = sValue + "%";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    

                    //dtB = pds.Tables["DataBlock4"];
                    //sPrefix = "B4";
                    //if (dtB != null)
                    //{
                    //    if (dtB.Rows.Count < 1) dtB.Rows.Add();
                    //    DataRow dr = dtB.Rows[0];

                    //    foreach (DataColumn col in dtB.Columns)
                    //    {
                    //        sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //        sValue = dr[col] + "";
                    //        if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //        if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                    //        rUtil.ReplaceTables(lstTable, sKey, sValue);
                    //    }
                    //}

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

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 2), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        double db6ObjSelfBearAmt = 0; //자기부담금
                        string tmp = "";

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjSelfBearAmt") db6ObjSelfBearAmt += Utils.ToDouble(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                        rUtil.ReplaceTables(lstTable, "@db6ObjSelfBearAmt@", Utils.AddComma(db6ObjSelfBearAmt));

                        if (Utils.ConvertToString(dr["InsurObjDvs"]) != "")
                        {
                            tmp += "\n" + dr["InsurObjDvs"] + "/" + dr["ObjStrt"];
                        }
                        rUtil.ReplaceTables(lstTable, "@db6ObjStrtRmk@", tmp);
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
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
                                //DataRow dr = dtB.Rows[0];

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "VitmNm");
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
                                    if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "VitmSubSeq") {
                                        if (dtB.Rows.Count == 1)
                                        {
                                            sValue = "";
                                        }
                                    }
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow((i + 1) * 2 + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow((i + 1) * 2 + 1), sKey, sValue);

                                }
                            }
                        }
                    }
                    int B12Count = dtB.Rows.Count;
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "MedHstr");
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
                                    if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "VitmSubSeq")
                                    {
                                        if (B12Count == 1)
                                        {
                                            sValue = "";
                                        }
                                        else
                                        {
                                            sValue += "-";
                                        }
                                    }
                                    if (col.ColumnName == "CureSeq")
                                    {
                                        if (dtB.Rows.Count == 1)
                                        {
                                            sValue = "";
                                        }
                                    }
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 1), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 2), sKey, sValue);
                                }
                            }
                        }
                    }

                    double db13ExpsLosAmt91 = 0;   //소계
                    double db13ExpsLosAmt7 = 0;    //피해자과실
                    double db13ExpsLosAmtCha = 0;  //과실상계후 금액

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        //1.치료비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        TableRow oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt1@", sExpsCmnt);
                        }

                        //2.휴업손해
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt2@", sExpsCmnt);
                        }

                        //3.상실수익
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt3@", sExpsCmnt);
                        }

                        //4.향후치료비
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        TableRow oRowBase = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt4@");
                        dReq = 0;
                        dAmt = 0;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableH, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsSubHed4@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt4@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt4@", drs[i]["ExpsCmnt"] + "");
                            }
                        }

                        //5.개호비
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt5@", sExpsCmnt);
                        }

                        //6.기타손해
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        oRowBase = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt6@");
                        dReq = 0;
                        dAmt = 0;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableH, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsSubHed6@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt6@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt6@", drs[i]["ExpsCmnt"] + "");
                            }
                        }

                        //91.소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt91@", Utils.AddComma(dAmt));
                        }
                        db13ExpsLosAmt91 = dAmt;

                        //7.과실부담금
                        drs = dtB?.Select("ExpsGrp = 7");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt7@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt7@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt7@", sExpsCmnt);
                        }
                        db13ExpsLosAmt7 = dAmt;

                        //8.위자료
                        drs = dtB?.Select("ExpsGrp = 8");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt8@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt8@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt8@", sExpsCmnt);
                        }

                        //9.자기부담금
                        drs = dtB?.Select("ExpsGrp = 9");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt9@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt9@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt9@", sExpsCmnt);
                        }

                        //92.합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt92@", Utils.AddComma(dAmt));
                        }

                        //93.예상지급보험금
                        drs = dtB?.Select("ExpsGrp = 93");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt93@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt93@", Utils.AddComma(dAmt));
                        }
                    }
                    db13ExpsLosAmtCha = db13ExpsLosAmt91 - db13ExpsLosAmt7;
                    rUtil.ReplaceTables(lstTable, "@db13ExpsLosAmtCha@", Utils.AddComma(db13ExpsLosAmtCha));

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
