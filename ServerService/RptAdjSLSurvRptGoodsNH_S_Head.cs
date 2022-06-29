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
    public class RptAdjSLSurvRptGoodsNH_S_Head
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
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjPrsCndt@");
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtDtTm@");
                    Table oTbl사고사항 = rUtil.GetTable(lstTable, "@B1AcdtCaus@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                    Table oTbl보험금지급처 = oTblRow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    dtB = pds.Tables["DataBlock3"];
                    if (dtB != null)
                    {
                        //1,총괄표
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl총괄표, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName("B15", "AcdtPictImage");
                        Table oTableA = rUtil.GetSubTable(oTbl사고사항, sKey);
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                //rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                //rUtil.TableAddRow(oTbl손해상황, 2, 1);
                                rUtil.TableAddRow(oTableA, 0, 1);
                                rUtil.TableAddRow(oTableA, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
                    if (dtB != null)
                    {
                        if (oTbl보험금지급처 != null)
                        {
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@db3ObjInsurRegsAmt@");
                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@B2InsurPrdt@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B6OthInsurCo@");
                    Table oTbl피보험자관련사항 = rUtil.GetTable(lstTable, "@B2IsrdRentCtrt@");
                    Table oTbl목적물현황 = rUtil.GetTable(lstTable, "@B3ObjPrsCndt@");
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtCaus@");
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B17InsurGivObj@");
                    Table oTbl보험금지급처 = oTblRow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    Table oTableB = rUtil.GetTable(lstTable, "@db3ObjSelfBearAmt@");                //자기부담금
                    Table oTableC = rUtil.GetSubTable(oTbl보험계약사항, "@B3ObjSymb@");             //2.보험계약사항 - 보험목적물 및 보험가입금액
                    Table oTableD = rUtil.GetSubTable(oTbl피보험자관련사항, "@B4InsurObjNm@");      //3.일반사항 - 가.피보험자 관련사항
                    Table oTableE = rUtil.GetSubTable(oTbl건물현황및배치도, "@B4ObjSymb_12@");      //3.일반사항 - 다.건물현황 및 배치도
                    Table oTableF = rUtil.GetSubTable(oTbl기계배치도, "@B4ObjSymb_13@");            //3.일반사항 - 라.기계배치도
                    Table oTableG = rUtil.GetSubTable(oTbl건물현황및배치도, "@B7AcdtPictImage@");   //3.일반사항 - 다.건물현황 및 배치도
                    Table oTableH = rUtil.GetSubTable(oTbl건물현황및배치도, "@B12AcdtPictImage@");
                    Table oTableI = rUtil.GetSubTable(oTbl기계배치도, "@B13AcdtPictImage@");
                    Table oTableJ = rUtil.GetSubTable(oTbl사고내용, "@B15AcdtPictImage@");

                    sKey = "@B1AcdtCaus@";
                    oTblA = rUtil.GetTable(lstTable, sKey);
                    sKey = "@B15AcdtPictImage1@";
                    oTblRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    Table oTableA = oTblRow?.GetCell(1).Elements<Table>().FirstOrDefault();

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
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "SealPhoto" || col.ColumnName == "ChrgAdjPhoto")
                            //{
                            //    try
                            //    {
                            //        Image img = Utils.stringToImage(sValue);
                            //        rUtil.ReplaceInternalImage(sKey, img);
                            //    }
                            //    catch { }
                            //    continue;
                            //}
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
                                if (col.ColumnName == "ObjInsurRegsAmt")
                                {
                                    db3ObjInsurRegsAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjInsValueTot")
                                {
                                    db3ObjInsValueTot += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjLosAmt")
                                {
                                    db3ObjLosAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjRmnAmt")
                                {
                                    db3ObjRmnAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjTotAmt")
                                {
                                    db3ObjTotAmt += Utils.ToDouble(sValue);
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

                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                //rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);          //2.보험계약사항 - 보험목적물 및 보험가입금액
                                //rUtil.ReplaceTableRow(oTbl목적물현황.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsurRegsAmt@", Utils.AddComma(db3ObjInsurRegsAmt));//보험가입금액
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjInsValueTot@", Utils.AddComma(db3ObjInsValueTot));//보험가액
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt));//손해액
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjRmnAmt@", Utils.AddComma(db3ObjRmnAmt));//잔존물
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjTotAmt@", Utils.AddComma(db3ObjTotAmt));//순손해액
                    //rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));//보험가입금액
                    rUtil.ReplaceTable(oTableB, "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));//자기부담금
                    rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt));//지급보험금

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
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
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableJ.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
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
                                if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue); //지급보험금
                                rUtil.ReplaceTableRow(oTbl보험금지급처.GetRow(i + 1), sKey, sValue);
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
