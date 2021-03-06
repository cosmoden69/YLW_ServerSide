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
    public class RptAdjSLSurvRptLiabilityDB_Head
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B2AcdtPictImage@");
                    Table oTbl현장배치도 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");
                    Table oTbl피해상황 = rUtil.GetTable(lstTable, "@B11DmobNm@");

                    dtB = pds.Tables["DataBlock2"];
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


                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        //3.현장배치도
                        if (oTbl현장배치도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl현장배치도, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    if (dtB != null)
                    {
                        //2.피해상황
                        if (oTbl피해상황 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl피해상황, 1, dtB.Rows.Count - 1);
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B2AcdtPictImage@");
                    Table oTbl현장배치도 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");
                    Table oTbl피해상황 = rUtil.GetTable(lstTable, "@B11DmobNm@");


                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        if (!dtB.Columns.Contains("AcdtDtTm")) dtB.Columns.Add("AcdtDtTm");
                        dr["AcdtDtTm"] = Utils.DateConv(dr["AcdtDt"] + "", ".") + " " + Utils.TimeConv(dr["AcdtTm"] + "", ":", "SHORT");
                        //if (!dtB.Columns.Contains("AcdtJurdPolcText")) dtB.Columns.Add("AcdtJurdPolcText");
                        //{
                        //    var AcdtJurdPolc = dr["AcdtJurdPolc"] + "";
                        //    var AcdtJurdPolcOpni = dr["AcdtJurdPolcOpni"] + "";
                        //    var AcdtJurdFire = dr["AcdtJurdFire"] + "";
                        //    var AcdtJurdFireOpni = dr["AcdtJurdFireOpni"] + "";
                        //    var text = AcdtJurdPolc + "\n" + AcdtJurdPolcOpni + "\n" + AcdtJurdFire + "\n" + AcdtJurdFireOpni;
                        //    dr["AcdtJurdPolcText"] = text;
                        //}

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
                            //1.총괄표
                            //if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "InsurRegsAmt2") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotReq") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AgrmAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSelfBearAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "SelfBearAmt2") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurAmt") sValue = Utils.AddComma(sValue);
                            //보험계약사항
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);

                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
                            if (col.ColumnName == "LeadAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "ChrgAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "BistLicSerl")
                            {
                                if (sValue != "") sValue = "보 조 인 등록번호 : 제 " + sValue + " 호";
                            }
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
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
                                        rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2200000L, 1500000L);
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

                    //dtB = pds.Tables["DataBlock2"];
                    //sPrefix = "B2";
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

                    //        //sKey = rUtil.GetFieldName(sPrefix, "ObjSymb");
                    //        //sValue = dr["ObjSymb"] + "";
                    //        //rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                    //        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                    //        sValue = dr["AcdtPictImage"] + "";
                    //        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                    //        try
                    //        {
                    //            Image img = Utils.stringToImage(sValue);
                    //            rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2200000L, 1500000L);
                    //        }
                    //        catch { }

                    //        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                    //        sValue = dr["AcdtPictCnts"] + "";
                    //        TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                    //        rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);

                    //    }

                    //}

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
                            //if (col.ColumnName == "CtrtDt") sValue = Utils.DateConv(sValue, ".");
                            //if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateConv(sValue, ".");
                            //if (col.ColumnName == "IsrdOpenDt") sValue = Utils.DateConv(sValue, ".");
                            //if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "IsrdEmpCnt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl현장배치도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl현장배치도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl현장배치도.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
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
                                rUtil.ReplaceTableRow(oTbl피해상황.GetRow(i + 1), sKey, sValue);
                                ////rUtil.ReplaceTableRow(oTbl보험계약사항.GetRow(i + 1), sKey, sValue);
                                ////rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);          //2.보험계약사항 - 보험목적물 및 보험가입금액
                                //rUtil.ReplaceTableRow(oTbl피해상황.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjInsurRegsAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
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
