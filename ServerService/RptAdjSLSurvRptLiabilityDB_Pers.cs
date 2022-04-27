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
    class RptAdjSLSurvRptLiabilityDB_Pers
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    //변수가 replace 되기 전에 테이블을 찾아 놓는다

                    Table oTableC = rUtil.GetTable(lstTable, "@B15ExpsLosAmt92@");

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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSubTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DIGivInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtCausCatg2Nm") sValue = Utils.AddComma(sValue);
                            
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    var db9DmgText = "";

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;
                            
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";

                                if (col.ColumnName == "VitmSubSeq")
                                {
                                    //var DgnsNm = dr["DgnsNm"] + ""; if (DgnsNm != null) { }
                                    //var VstHosp =  dr["VstHosp"] + "";
                                    //var CureFrDt = dr["CureFrDt"] + "";
                                    //var CureToDt = dr["CureToDt"] + "";
                                    //var DoctDgns = dr["DoctDgns"] + "";
                                    //var MedHstr = dr["MedHstr"] + "";
                                    //var CureCnts = dr["CureCnts"] + "";
                                    //var CureMjrCnts = dr["CureMjrCnts"] + "";
                                    //db9DmgText += "\n" + DgnsNm + "\n" + VstHosp + "\n" + Utils.DateFormat(CureFrDt, "yyyy.MM.dd") + " ~ " + Utils.DateFormat(CureToDt, "yyyy.MM.dd") + "\n"
                                    //              + DoctDgns + "\n" + MedHstr + "\n" + CureCnts + ", " + CureMjrCnts + "\n";

                                    var DgnsNm = dr["DgnsNm"] + ""; 
                                    var VstHosp = dr["VstHosp"] + ""; 
                                    var CureFrDt = dr["CureFrDt"] + ""; 
                                    var CureToDt = dr["CureToDt"] + ""; 
                                    var DoctDgns = dr["DoctDgns"] + ""; 
                                    var MedHstr = dr["MedHstr"] + ""; 
                                    var CureCnts = dr["CureCnts"] + ""; 
                                    var CureMjrCnts = dr["CureMjrCnts"] + ""; 

                                    if (!(DgnsNm == null) && !(DgnsNm == "")) {db9DmgText += "\n" + DgnsNm; }
                                    if (!(VstHosp == null) && !(VstHosp == "")) { db9DmgText += "\n" + VstHosp; }
                                    if (!(CureFrDt == null) && !(CureFrDt == "")) { db9DmgText += "\n" + Utils.DateFormat(CureFrDt, "yyyy.MM.dd"); }
                                    if (!(CureToDt == null) && !(CureToDt == "")) { db9DmgText += " ~ " + Utils.DateFormat(CureToDt, "yyyy.MM.dd"); }
                                    if (!(DoctDgns == null) && !(DoctDgns == "")) { db9DmgText += "\n" + DoctDgns; }
                                    if (!(MedHstr == null) && !(MedHstr == "")) { db9DmgText += "\n" + MedHstr; }
                                    if (!(CureCnts == null) && !(CureCnts == "")) { db9DmgText += "\n" + CureCnts; }
                                    if (!(CureMjrCnts == null) && !(CureMjrCnts == "")) { db9DmgText += ", " + CureMjrCnts + "\n"; }
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                            }
                        }
                    }
                    rUtil.ReplaceTextAllParagraph(doc, "@db9DmgText@", db9DmgText); //가.평가기준 - 6) 피해정도

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        //1.치료비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        TableRow oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq1@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt1@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss1@", sExpsBss);
                        }

                        //2.휴업손해
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq2@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt2@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss2@", sExpsBss);
                        }

                        //3.상실수익
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq3@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt3@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss3@", sExpsBss);
                        }

                        //4.향후치료비
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
                        TableRow oRowBase = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq4@");
                        int rIdx1 = -1;
                        int rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableC, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsSubHed4@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq4@", Utils.AddComma(drs[i]["ExpsLosReq"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt4@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt4@", drs[i]["ExpsCmnt"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsBss4@", drs[i]["ExpsBss"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTableC, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTableC, oRow);
                        }
                        rUtil.TableMergeCellsV(oTableC, 0, rIdx1, rIdx2);

                        //5.개호비
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq5@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt5@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss5@", sExpsBss);
                        }

                        //6.기타손해
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
                        oRowBase = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq6@");
                        rIdx1 = -1;
                        rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableC, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsSubHed6@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq6@", Utils.AddComma(drs[i]["ExpsLosReq"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt6@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt6@", drs[i]["ExpsCmnt"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B15ExpsBss6@", drs[i]["ExpsBss"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTableC, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTableC, oRow);
                        }
                        rUtil.TableMergeCellsV(oTableC, 0, rIdx1, rIdx2);

                        //91.소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq91@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt91@", Utils.AddComma(dAmt));
                        }

                        //7.과실부담금
                        drs = dtB?.Select("ExpsGrp = 7");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq7@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq7@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt7@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt7@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss7@", sExpsBss);
                        }

                        //8.위자료
                        drs = dtB?.Select("ExpsGrp = 8");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq8@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq8@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt8@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt8@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss8@", sExpsBss);
                        }

                        //92.합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq92@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt92@", Utils.AddComma(dAmt));
                        }

                        //9.자기부담금
                        drs = dtB?.Select("ExpsGrp = 9");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq9@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq9@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt9@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsCmnt9@", sExpsCmnt);
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsBss9@", sExpsBss);
                        }

                        //93.예상지급보험금
                        drs = dtB?.Select("ExpsGrp = 93");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock15"].Rows.Add() };
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
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B15ExpsLosReq93@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosReq93@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B15ExpsLosAmt93@", Utils.AddComma(dAmt));
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
