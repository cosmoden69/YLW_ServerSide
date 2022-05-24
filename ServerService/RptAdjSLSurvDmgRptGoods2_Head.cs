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
    public class RptAdjSLSurvDmgRptGoods2_Head
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@db3ObjInsurRegsAmt@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B2InsurPrdt@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B6OthInsurCo@");
                    Table oTbl피보험자관련사항 = rUtil.GetTable(lstTable, "@B2IsrdRentCtrt@");
                    Table oTbl피해자관련사항 = rUtil.GetTable(lstTable, "@B1VitmRegno@");
                    Table oTbl현장도면 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtDtTm@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                    Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    dtB = pds.Tables["DataBlock3"];
                    if (dtB != null)
                    {
                        //1,총괄표
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl총괄표, 1, dtB.Rows.Count - 1);
                        }

                        //2.보험계약사항 - 담보위험
                        Table oTableC = rUtil.GetSubTable(oTbl보험계약사항, "@B3ObjSymb@");
                        if (oTableC != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTableC, 1, dtB.Rows.Count - 1);
                        }
                    }

                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1");
                    Table oTableF = rUtil.GetSubTable(oTbl현장도면, "@B4ObjSymb_12@");
                    if (oTableF != null)
                    {
                        if (drs != null && drs.Length > 0)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTableF, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    if (dtB != null)
                    {
                        //2.보험계약사항 - 타보험 계약사항
                        if (oTbl타보험계약사항 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl타보험계약사항, 2, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        //3.일반사항 - 다.현장도면
                        Table oTableG = rUtil.GetSubTable(oTbl현장도면, "@B7AcdtPictImage@");
                        if (oTableG != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRows(oTableG, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    if (dtB != null)
                    {
                        Table oTableH = rUtil.GetSubTable(oTbl현장도면, "@B12AcdtPictImage@");
                        if (oTableH != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 2) / 3.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableH, 0, 1);
                                rUtil.TableAddRow(oTableH, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock14"];
                    if (dtB != null)
                    {
                        Table oTableJ = rUtil.GetSubTable(oTbl사고내용, "@B14AcdtPictImage@");
                        if (oTableJ != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableJ, 0, 1);
                                rUtil.TableAddRow(oTableJ, 1, 1);
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
                    Table oTbl표지 = rUtil.GetTable(lstTable, "@B1LeadAdjuster@");
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@db3ObjInsurRegsAmt@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B2InsurPrdt@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B6OthInsurCo@");
                    Table oTbl피보험자관련사항 = rUtil.GetTable(lstTable, "@B2IsrdRentCtrt@");
                    Table oTbl피해자관련사항 = rUtil.GetTable(lstTable, "@B1VitmRegno@");
                    Table oTbl현장도면 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtDtTm@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                    Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    Table oTableC = rUtil.GetSubTable(oTbl보험계약사항, "@B3ObjSymb@");             //2.보험계약사항 - 담보위험
                    Table oTableF = rUtil.GetSubTable(oTbl현장도면, "@B4ObjSymb_12@");              //3.일반사항
                    Table oTableG = rUtil.GetSubTable(oTbl현장도면, "@B7AcdtPictImage@");           //3.일반사항
                    Table oTableH = rUtil.GetSubTable(oTbl현장도면, "@B12AcdtPictImage@");
                    Table oTableJ = rUtil.GetSubTable(oTbl사고내용, "@B14AcdtPictImage@");

                    var db1SurvAsgnEmpManRegNo = ""; //조사자 손해사정등록번호
                    var db1SurvAsgnEmpAssRegNo = ""; //조사자 보조인 등록번호
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
                            if (col.ColumnName == "LeadAdjManRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "ChrgAdjManRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "SurvAsgnEmpManRegNo")
                            {
                                if (sValue != "") db1SurvAsgnEmpManRegNo = sValue;
                            }
                            if (col.ColumnName == "SurvAsgnEmpAssRegNo")
                            {
                                if (sValue != "") db1SurvAsgnEmpAssRegNo = sValue;
                            }
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    if (db1SurvAsgnEmpManRegNo == "")
                    {
                        if (db1SurvAsgnEmpAssRegNo != "")
                        {
                            db1SurvAsgnEmpAssRegNo = "보조인 등록번호 : 제" + db1SurvAsgnEmpAssRegNo + "호";
                        }
                        rUtil.ReplaceTable(oTbl표지, "@db1SurvAsgnEmpRegNo@", db1SurvAsgnEmpAssRegNo);
                    }
                    else
                    {
                        db1SurvAsgnEmpManRegNo = "손해사정등록번호 : 제" + db1SurvAsgnEmpManRegNo + "호";
                        rUtil.ReplaceTable(oTbl표지, "@db1SurvAsgnEmpRegNo@", db1SurvAsgnEmpManRegNo);
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

                    double db3ObjInsurRegsAmt = 0;
                    double db3ObjLosAmt = 0;
                    double db3AgrmAmt = 0;
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
                                if (col.ColumnName == "ObjInsurRegsAmt")
                                {
                                    db3ObjInsurRegsAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjLosAmt")
                                {
                                    db3ObjLosAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "AgrmAmt")
                                {
                                    db3AgrmAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjSelfBearAmt")
                                {
                                    db3ObjSelfBearAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjGivInsurAmt")
                                {
                                    db3ObjGivInsurAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjInsValueTot") sValue = Utils.AddComma(sValue);
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);          //2.보험계약사항 - 담보위험
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3AgrmAmt@", Utils.AddComma(db3AgrmAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt));
                    rUtil.ReplaceTable(oTbl보험계약사항, "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));

                    double db4ObjArea = 0;

                    //건물구조 및 면적
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1");
                    sPrefix = "B4";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTableF != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = "@B4" + col.ColumnName + "_12@";
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjArea")
                                    {
                                        db4ObjArea += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTableF.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (oTbl타보험계약사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (i + 1) * 2;
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "OthSelfBearAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTableG != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableG.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableG.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (oTableH != null)
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
                                TableRow xrow1 = oTableH.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2000000L, 1400000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableH.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (oTableJ != null)
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
                                TableRow xrow1 = oTableJ.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableJ.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
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

                            sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                            sValue = dr["AcdtPictImage"] + "";
                            rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                            try
                            {
                                Image img = Utils.stringToImage(sValue);
                                rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2700000L, 2000000L);
                            }
                            catch { }

                            sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                            sValue = dr["AcdtPictCnts"] + "";
                            TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                            rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
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
