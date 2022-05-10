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
    public class RptAdjSLSurvRptGoodsNH_Head
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B3ObjSymb@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B3EvatRsltStrt@");
                    Table oTbl피보험자개요 = rUtil.GetTable(lstTable, "@B2MonSellAmt@");
                    Table oTbl건물구조및면적 = rUtil.GetSubTable(oTbl피보험자개요, "@B4ObjStrt_12@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B6OthInsurCo@");
                    Table oTbl피보험자관련사항 = rUtil.GetTable(lstTable, "@B2IsrdRentCtrt@");
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjPrsCndt@");
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl사고현장사진 = rUtil.GetTable(lstTable, "@B14AcdtPictImage@");
                    Table oTbl건물현황및개요 = rUtil.GetTable(lstTable, "@B4ObjSymb_12@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl기계범례전체 = rUtil.GetTable(lstTable, "<기계범례>");
                    Table oTbl기계범례 = rUtil.GetSubTable(oTbl기계범례전체, "@B4ObjSymb_13@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtDtTm@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");

                    dtB = pds.Tables["DataBlock3"];
                    if (dtB != null)
                    {
                        //1.총괄표
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl총괄표, 1, dtB.Rows.Count - 1);
                        }

                        //2.보험계약사항 - 보험목적물 및 보험가입금액
                        if (oTbl보험계약사항 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl보험계약사항, 1, dtB.Rows.Count - 1);
                        }

                        //3.일반사항 - 나.목적물현황
                        if (oTbl목적물현황 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl목적물현황, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //건물구조 및 면적
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1 OR ObjCatgCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        //3.일반사항 - 다.건물현황 및 배치도
                        sKey = rUtil.GetFieldName(sPrefix, "B4ObjSymb_12");
                        if (oTbl건물구조및면적 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl건물구조및면적, 1, drs.Length - 1);
                        }
                        if (oTbl건물현황및개요 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl건물현황및개요, 1, drs.Length - 1);
                        }
                    }

                    //3.일반사항 - 라.기계배치도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 3 OR ObjCatgCd % 10 = 4");
                    if (drs != null && drs.Length > 0)
                    {
                        //3.일반사항 - 라.기계배치도
                        if (oTbl기계범례 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl기계범례, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "B6OthInsurCo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }

                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "B7AcdtPictImage");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTable, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "B13AcdtPictImage");
                        if (oTbl기계배치도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl기계배치도, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock14"];
                    if (dtB != null)
                    {
                        if (oTbl사고현장사진 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 2) / 3.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl사고현장사진, 0, 1);
                                rUtil.TableAddRow(oTbl사고현장사진, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                rUtil.TableAddRow(oTbl손해상황, 2, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock17"];
                    if (dtB != null)
                    {
                        Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                        TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                        Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();
                        if (oTbl보험금지급처 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl보험금지급처, 1, dtB.Rows.Count - 1);
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B3ObjSymb@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B3EvatRsltStrt@");
                    Table oTbl피보험자개요 = rUtil.GetTable(lstTable, "@B2MonSellAmt@");
                    Table oTbl건물구조및면적 = rUtil.GetSubTable(oTbl피보험자개요, "@B4ObjStrt_12@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B6OthInsurCo@");
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjPrsCndt@");
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl사고현장사진 = rUtil.GetTable(lstTable, "@B14AcdtPictImage@");
                    Table oTbl건물현황및개요 = rUtil.GetTable(lstTable, "@B4ObjSymb_12@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl기계범례전체 = rUtil.GetTable(lstTable, "<기계범례>");
                    Table oTbl기계범례 = rUtil.GetSubTable(oTbl기계범례전체, "@B4ObjSymb_13@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                    Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        if (!dtB.Columns.Contains("AcdtDtTm")) dtB.Columns.Add("AcdtDtTm");
                        dr["AcdtDtTm"] = Utils.DateConv(dr["AcdtDt"] + "", ".") + " " + Utils.TimeConv(dr["AcdtTm"] + "", ":", "SHORT");
                        if (!dtB.Columns.Contains("AcdtJurdPolcText")) dtB.Columns.Add("AcdtJurdPolcText");
                        {
                            var AcdtJurdPolc = dr["AcdtJurdPolc"] + "";
                            var AcdtJurdPolcOpni = dr["AcdtJurdPolcOpni"] + "";
                            var AcdtJurdFire = dr["AcdtJurdFire"] + "";
                            var AcdtJurdFireOpni = dr["AcdtJurdFireOpni"] + "";
                            var text = AcdtJurdPolc + "\n" + AcdtJurdPolcOpni + "\n" + AcdtJurdFire + "\n" + AcdtJurdFireOpni;
                            dr["AcdtJurdPolcText"] = text;
                        }

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
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH시 mm분경");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy. MM. dd.");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy. MM. dd.");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy. MM. dd.");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
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
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "IsrdOpenDt")
                            {
                                sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년 " + Utils.Mid(sValue, 5, 6) + "월경");
                            }
                            if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "IsrdEmpCnt") sValue = (sValue == "" ? "" : Utils.AddComma(sValue) + "명");
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db3ObjInsurRegsAmt = 0;
                    double db3ObjInsValueTot = 0;
                    double db3ObjLosAmt = 0;
                    double db3ObjRmnAmt = 0;
                    double db3ObjTotAmt = 0;
                    double db3ObjSelfBearAmt = 0;
                    double db3ObjGivInsurAmt = 0;

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
                                if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                if (col.ColumnName == "ObjInsurRegsAmt")//보험가입금액
                                {
                                    db3ObjInsurRegsAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjInsValueTot")//보험가액
                                {
                                    db3ObjInsValueTot += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjLosAmt")//손해액
                                {
                                    db3ObjLosAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjRmnAmt")//잔존물
                                {
                                    db3ObjRmnAmt += Utils.ToDouble(sValue);
                                    sValue = (Utils.ToDouble(sValue) == 0 ? "-" : Utils.AddComma(sValue));
                                }
                                if (col.ColumnName == "ObjTotAmt")//순손해액
                                {
                                    db3ObjTotAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                //if (col.ColumnName == "ObjSelfBearAmt")
                                //{
                                //    db3ObjSelfBearAmt += Utils.ToDouble(sValue);
                                //    sValue = Utils.AddComma(sValue);
                                //}
                                if (col.ColumnName == "ObjGivInsurAmt")//지급보험금
                                {
                                    db3ObjGivInsurAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(i + 1), sKey, sValue);
                                //rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);          //2.보험계약사항 - 보험목적물 및 보험가입금액
                                rUtil.ReplaceTableRow(oTbl목적물현황.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt));
                    rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsValueTot@", Utils.AddComma(db3ObjInsValueTot));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjRmnAmt@", Utils.AddComma(db3ObjRmnAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjTotAmt@", Utils.AddComma(db3ObjTotAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt));

                    double db4ObjArea = 0;

                    //건물현황 및 배치도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1 OR ObjCatgCd % 10 = 2");
                    sPrefix = "B4";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl건물구조및면적 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = "@B4" + col.ColumnName + "_12@";
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                    if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "ObjBuyDt") sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년" + Utils.Mid(sValue, 5, 6) + "월");
                                    rUtil.ReplaceTableRow(oTbl건물구조및면적.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                        if (oTbl건물현황및개요 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = "@B4" + col.ColumnName + "_12@";
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                    if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "ObjBuyDt") sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년" + Utils.Mid(sValue, 5, 6) + "월");
                                    rUtil.ReplaceTableRow(oTbl건물현황및개요.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    if (dtB == null || dtB.Rows.Count < 1)
                    {
                        //기계 배치도 사진 없으면 아래 기계범례 삭제
                        Table tbl = rUtil.GetTable(lstTable, "<기계범례>");
                        if (tbl != null) tbl.Remove();
                    }
                    else
                    {
                        //기계 배치도
                        drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 3 OR ObjCatgCd % 10 = 4");
                        sPrefix = "B4";
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                        if (drs != null && drs.Length > 0)
                        {
                            if (oTbl기계범례 != null)
                            {
                                for (int i = 0; i < drs.Length; i++)
                                {
                                    DataRow dr = drs[i];
                                    foreach (DataColumn col in dr.Table.Columns)
                                    {
                                        sKey = "@B4" + col.ColumnName + "_13@";
                                        sValue = dr[col] + "";
                                        if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                        if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                        if (col.ColumnName == "ObjBuyDt") sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년" + Utils.Mid(sValue, 5, 6) + "월");
                                        rUtil.ReplaceTableRow(oTbl기계범례.GetRow(i + 1), sKey, sValue);
                                    }
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "OthInsurCo");
                        if (oTbl타보험계약사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();

                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];

                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateConv(sValue, ".");
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl타보험계약사항.Remove();
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl건물현황및배치도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl건물현황및배치도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl건물현황및배치도.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (oTbl사고현장사진 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 3 != 0)
                            {
                                if (dtB.Rows.Count % 3 >= 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                                if (dtB.Rows.Count % 3 >= 2) dtB.Rows.Add();  //세번째 칸을 클리어 해주기 위해서 추가
                            }
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 3.0) * 2;
                                int rmdr = i % 3;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl사고현장사진.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2000000L, 1400000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl사고현장사진.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        if (oTbl기계배치도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl기계배치도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl기계배치도.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
                        { 
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                                int rmdr = i % 2 + 1;

                                TableRow xrow1 = oTbl손해상황.GetRow(rnum);

                                sKey = rUtil.GetFieldName(sPrefix, "ObjNm");
                                sValue = dr["ObjNm"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "ObjSymb");
                                sValue = dr["ObjSymb"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2200000L, 1500000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictRmk");
                                sValue = dr["AcdtPictRmk"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
                    if (dtB != null)
                    {
                        if (oTbl보험금지급처 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl보험금지급처.GetRow(i + 1), sKey, sValue);
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
