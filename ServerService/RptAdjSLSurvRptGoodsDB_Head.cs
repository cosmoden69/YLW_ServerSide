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
    public class RptAdjSLSurvRptGoodsDB_Head
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B3ObjInsurRegsAmt@");
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjSymb@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B4ObjInsureValue@");
                    Table oTbl건물내역 = rUtil.GetTable(lstTable, "@B4ObjSymb_12@");
                    Table oTbl기계범례 = rUtil.GetTable(lstTable, "@B4ObjSymb_13@");
                    Table oTbl건물현황도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
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
                        
                        //4.일반사항 - 나.목적물현황
                        if (oTbl목적물현황 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl목적물현황, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    if (dtB != null)
                    {
                        //2.보험계약사항 - 아.보험목적물 및 해당 보험가입금액
                        if (oTbl보험계약사항 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl보험계약사항, 1, dtB.Rows.Count - 1);
                        }
                    }


                    //4.일반사항 - 다.건물현황도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1 OR ObjCatgCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl건물내역 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl건물내역, 1, drs.Length - 1);
                        }
                    }

                    //4.일반사항 - 라.기계배치도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 3 OR ObjCatgCd % 10 = 4");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl기계범례 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl기계범례, 1, drs.Length - 1);
                        }
                    }
                    
                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        //4.일반사항 - 다.건물현황도
                        if (oTbl건물현황도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl건물현황도, 0, 2, dtB.Rows.Count - 1);
                        }
                    }


                    dtB = pds.Tables["DataBlock13"];
                    if (dtB != null)
                    {
                        //4.일반사항 - 라.기계배치도
                        if (oTbl기계배치도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl기계배치도, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    DataView view = dtB.DefaultView;
                    DataTable dtN = view.ToTable(true, "ObjSeq");
                    if (dtN != null && oTbl손해상황 != null)
                    {
                        double cnt = 0;
                        foreach (DataRow drN in dtN.Rows)
                        {
                            DataTable dtNS = dtB.Select("ObjSeq= '" + drN["ObjSeq"] + "' ")?.CopyToDataTable();
                            //테이블의 끝에 추가
                            cnt += Math.Truncate((dtNS.Rows.Count + 1) / 2.0);
                        }
                        for (int i = 1; i < cnt; i++)
                        {
                            rUtil.TableAddRow(oTbl손해상황, 1, 1);
                            rUtil.TableAddRow(oTbl손해상황, 2, 1);
                        }
                    }

                    //dtB = pds.Tables["DataBlock15"];
                    //sPrefix = "B15";
                    //DataRow[] drs1 = pds.Tables["DataBlock15"]?.Select("ObjSeq  = 1 ");
                    //DataRow[] drs2 = pds.Tables["DataBlock15"]?.Select("ObjSeq  = 2 ");
                    //DataRow[] drs3 = pds.Tables["DataBlock15"]?.Select("ObjSeq  = 3 ");
                    //DataRow[] drs4 = pds.Tables["DataBlock15"]?.Select("ObjSeq  = 4 ");

                    ////Count
                    //int drs1Count = drs1.Count();
                    //int drs2Count = drs2.Count();
                    //int drs3Count = drs3.Count();
                    //int drs4Count = drs4.Count();

                    //// 늘어나야할 행의 수
                    //int drs1Addrow = 0; 
                    //int drs2Addrow = 0; 
                    //int drs3Addrow = 0;
                    //int drs4Addrow = 0;
                    ////행이 0 일때
                    //if (drs1Count == 0) drs1Addrow = 0;
                    //if (drs2Count == 0) drs2Addrow = 0;
                    //if (drs3Count == 0) drs3Addrow = 0;
                    //if (drs4Count == 0) drs4Addrow = 0;
                    ////짝수 일때
                    //if (drs1Count % 2 == 0) { drs1Addrow = drs1Count / 2; }
                    //if (drs2Count % 2 == 0) { drs2Addrow = drs2Count / 2; }
                    //if (drs3Count % 2 == 0) { drs3Addrow = drs3Count / 2; }
                    //if (drs4Count % 2 == 0) { drs4Addrow = drs4Count / 2; }
                    ////홀수
                    //if (drs1Count % 2 == 1) { drs1Addrow = (drs1Count + 1) / 2; }
                    //if (drs2Count % 2 == 1) { drs2Addrow = (drs2Count + 1) / 2; }
                    //if (drs3Count % 2 == 1) { drs3Addrow = (drs3Count + 1) / 2; }
                    //if (drs4Count % 2 == 1) { drs4Addrow = (drs4Count + 1) / 2; }

                    //if (dtB != null)
                    //{
                    //    if (oTbl손해상황 != null)
                    //    {
                    //        //테이블의 끝에 추가
                    //        //ObjSeq  = 1
                    //        for (int i = 1; i < drs1Addrow; i++)
                    //        {
                    //            rUtil.TableAddRow(oTbl손해상황, 1, 1);
                    //            rUtil.TableAddRow(oTbl손해상황, 2, 1);
                    //        }
                    //        //ObjSeq  = 2
                    //        for (int i = 0; i < drs2Addrow; i++)
                    //        {
                    //            rUtil.TableAddRow(oTbl손해상황, 1, 1);
                    //            rUtil.TableAddRow(oTbl손해상황, 2, 1);
                    //        }
                    //        //ObjSeq  = 3
                    //        for (int i = 0; i < drs3Addrow; i++)
                    //        {
                    //            rUtil.TableAddRow(oTbl손해상황, 1, 1);
                    //            rUtil.TableAddRow(oTbl손해상황, 2, 1);
                    //        }
                    //        //ObjSeq  = 4
                    //        for (int i = 0; i < drs4Addrow; i++)
                    //        {
                    //            rUtil.TableAddRow(oTbl손해상황, 1, 1);
                    //            rUtil.TableAddRow(oTbl손해상황, 2, 1);
                    //        }
                    //    }
                    //}

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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B3ObjInsurRegsAmt@");
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjSymb@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B4ObjInsureValue@");
                    Table oTbl건물내역 = rUtil.GetTable(lstTable, "@B4ObjSymb_12@");
                    Table oTbl기계범례 = rUtil.GetTable(lstTable, "@B4ObjSymb_13@");
                    Table oTbl건물현황도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    

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
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SealPhoto" || col.ColumnName == "ChrgAdjPhoto")
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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "IsrdOpenDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "IsrdEmpCnt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db3ObjInsurRegsAmt = 0; //보험가입액 합계
                    double db3ReDeliValue = 0; //재조달가액 합계
                    double db3ObjInsValueTot = 0; //보험가액 합계
                    double db3ObjLosAmt = 0; //손해액 합계
                    double db3DefuctedValue = 0; //공제금액 합계
                    double db3ObjRmnAmt = 0; //잔존물 합계
                    double db3PureLosAmt = 0; //순손해액 합계
                    double db3ObjGivInsurAmt = 0; //지급보험금 합계

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
                                if(col.ColumnName == "ReDeliValue")//재조달가액
                                {
                                    db3ReDeliValue += Utils.ToDouble(sValue);
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
                                if (col.ColumnName == "DeductedValue")//공제금액
                                {
                                    db3DefuctedValue += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjRmnAmt")//잔존물
                                {
                                    db3ObjRmnAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "PureLosAmt")//순손해액
                                {
                                    db3PureLosAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjGivInsurAmt")//지급보험금
                                {
                                    db3ObjGivInsurAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                //rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(i + 1), sKey, sValue);
                                //rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);          //2.보험계약사항 - 보험목적물 및 보험가입금액
                                rUtil.ReplaceTableRow(oTbl목적물현황.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt)); //보험가입액 합계
                    //rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ReDeliValue@", Utils.AddComma(db3ReDeliValue)); //재조달가액 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsValueTot@", Utils.AddComma(db3ObjInsValueTot)); //보험가액 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt)); //손해액 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3DefuctedValue@", Utils.AddComma(db3DefuctedValue)); //공재금액 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjRmnAmt@", Utils.AddComma(db3ObjRmnAmt)); //잔존물 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3PureLosAmt@", Utils.AddComma(db3PureLosAmt)); //순손해액 합계
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt)); //지급보험금 합계



                    //2.보험계약사항 - 8.보험목적물 및 보험가입금액
                    double db4ObjInsureValue = 0; //해당보험금액 합계
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (oTbl보험계약사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                    
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjInsureValue") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "ObjInsureValue")//해당보험금액 합계
                                    {
                                        db4ObjInsureValue += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                        
                                    }
                                    rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(i + 1), sKey, sValue);
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTextAllParagraph(doc, "@db4ObjInsureValue@", Utils.AddComma(db4ObjInsureValue)); //사.보험금액(해당보험금액 합계)
                    rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(dtB.Rows.Count + 1), "@db4ObjInsureValue@", Utils.AddComma(db4ObjInsureValue)); //아.보험목적물 및 해당보험가입금액(해당보험금액 합계)

                    //건물내역
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1 OR ObjCatgCd % 10 = 2");
                    sPrefix = "B4";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl건물내역 != null)
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
                                    rUtil.ReplaceTableRow(oTbl건물내역.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    //기계범례
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
                                    rUtil.ReplaceTableRow(oTbl기계범례.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    var db6OthInsureText = "";

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
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

                                if (col.ColumnName == "OthInsurCo")
                                {
                                    var OthInsurCo = dr["OthInsurCo"] + "";
                                    var OthInsurPrdt = dr["OthInsurPrdt"] + "";
                                    var OthCtrtDt = dr["OthCtrtDt"] + "";
                                    var OthCtrtExprDt = dr["OthCtrtExprDt"] + "";
                                    db6OthInsureText += OthInsurCo + ", " + OthInsurPrdt + ", " + Utils.DateFormat(OthCtrtDt, "yyyy년 MM월 dd일") + " ~ " + Utils.DateFormat(OthCtrtExprDt, "yyyy년 MM월 dd일") + "\n";
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                
                            }
                        }
                    }
                    rUtil.ReplaceTextAllParagraph(doc, "@db6OthInsureText@", db6OthInsureText); //3.다른보험계약사항

                    
                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl건물현황도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl건물현황도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl건물현황도.GetRow(rnum + 1);
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
                    if (dtB.Rows.Count < 1)
                    {
                        dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTable(oTbl손해상황, sKey, sValue);
                        }
                    }
                    else
                    {
                        DataView view = dtB.DefaultView;
                        DataTable dtN = view.ToTable(true, "ObjSeq");
                        if (dtN != null && oTbl손해상황 != null)
                        {
                            int rnumAdd = 1;
                            foreach (DataRow drN in dtN.Rows)
                            {
                                DataTable dtNS = dtB.Select("ObjSeq= '" + drN["ObjSeq"] + "' ")?.CopyToDataTable();
                                if (dtNS.Rows.Count < 1) dtNS.Rows.Add();
                                if (dtNS.Rows.Count % 2 == 1) dtNS.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                                for (int i = 0; i < dtNS.Rows.Count; i++)
                                {
                                    DataRow dr = dtNS.Rows[i];
                                    int rnum = (int)Math.Truncate(i / 2.0) * 2 + rnumAdd;
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

                                    sKey = rUtil.GetFieldName(sPrefix, "ObjInsurRegsFg");
                                    sValue = dr["ObjInsurRegsFg"] + "";
                                    if (sValue == "1")
                                    {
                                        sValue = "보험가입";
                                    }
                                    else if (sValue == null || sValue == "")
                                    {
                                        sValue = "보험 미가입";
                                    }
                                    rUtil.SetText(xrow1.GetCell(3), sKey, sValue);
                                }
                                rnumAdd += (int)Math.Truncate(dtNS.Rows.Count / 2.0) * 2;
                            }
                        }
                    }

                    //drs = pds.Tables["DataBlock15"]?.Select("ObjSeq = 1");
                    //dtB = drs.CopyToDataTable();

                    ////dtB = pds.Tables["DataBlock15"];
                    //sPrefix = "B15";
                    //if (dtB != null)
                    //{
                    //    if (dtB.Rows.Count < 1) dtB.Rows.Add();
                    //    if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                    //    for (int i = 0; i < dtB.Rows.Count; i++)
                    //    {
                    //        DataRow dr = dtB.Rows[i];
                    //        int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                    //        int rmdr = i % 2 + 1;

                    //        TableRow xrow1 = oTbl손해상황.GetRow(rnum);

                    //        sKey = rUtil.GetFieldName(sPrefix, "ObjNm");
                    //        sValue = dr["ObjNm"] + "";
                    //        rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                    //        sKey = rUtil.GetFieldName(sPrefix, "ObjSymb");
                    //        sValue = dr["ObjSymb"] + "";
                    //        rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                    //        //sKey = rUtil.GetFieldName(sPrefix, "ObjInsurRegsFg");
                    //        //sValue = dr["ObjInsurRegsFg"] + "";
                    //        //rUtil.SetText(xrow1.GetCell(3), sKey, "");

                    //        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                    //        sValue = dr["AcdtPictImage"] + "";
                    //        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                    //        try
                    //        {
                    //            Image img = Utils.stringToImage(sValue);
                    //            rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2200000L, 1500000L);
                    //        }
                    //        catch { }

                    //        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                    //        sValue = dr["AcdtPictCnts"] + "";
                    //        TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                    //        rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                    //    }
                    //}

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
