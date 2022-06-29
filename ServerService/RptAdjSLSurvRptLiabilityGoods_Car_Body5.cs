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
    public class RptAdjSLSurvRptLiabilityGoods_Car_Body5
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile, int ObjSeq)
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
                    Table oTbl렌트비 = rUtil.GetTable(lstTable, "@B12CarTyp@");
                    Table oTbl교통비 = rUtil.GetTable(lstTable, "@B13CarTyp@");

                    dtB = pds.Tables["DataBlock12"];
                    drs = dtB?.Select("DmobSeq = " + ObjSeq); //목적물 별로 RePlace
                    var B12RowCnt = drs.Length;
                    if (dtB != null)
                    {
                        //피해물 파손부위확인
                        if (oTbl렌트비 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl렌트비, 2, 1, B12RowCnt - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    drs = dtB?.Select("DmobSeq = " + ObjSeq); //목적물 별로 RePlace
                    var B13RowCnt = drs.Length;
                    if (dtB != null)
                    {
                        //피해물 파손부위확인
                        if (oTbl교통비 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl교통비, 2, 3, B13RowCnt - 1);
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
                    Table oTbl평가기준 = rUtil.GetTable(lstTable, "@B10DmobSortNo@");
                    Table oTbl평가결과 = rUtil.GetTable(lstTable, "@B11DmobSortNo@");
                    Table oTbl렌트비 = rUtil.GetTable(lstTable, "@B12CarTyp@");
                    Table oTbl교통비 = rUtil.GetTable(lstTable, "@B13CarTyp@");

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
                            if (col.ColumnName == "InsurRegsAmt2") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AgrmAmt") sValue = Utils.AddComma(sValue);
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
                    

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    drs = dtB?.Select("DmobSeq = " + ObjSeq);
                    if (drs != null)
                    {
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock10"].Rows.Add() };
                        for (int i = 0; i < drs.Length; i++)
                        {
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = Utils.ConvertToString(drs[i][col] + "");

                                rUtil.ReplaceTable(oTbl평가기준, sKey, sValue);
                            }
                        }
                    }

                    
                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {   
                        //렌트비
                        drs = dtB?.Select("ExpsCd = 300263001 AND DmobSeq = " + ObjSeq); //렌트비
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };
                        TableRow oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm1@");
                        TableRow  oRowBase = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm1@");
                        int rIdx1 = -1;
                        int rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTbl평가결과, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTable(oTbl평가결과, "@B11DmobSortNo@", drs[i]["DmobSortNo"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ExpsGrpNm1@", drs[i]["ExpsGrpNm"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ReqAmt1@", Utils.AddComma(drs[i]["ReqAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11DoLosAmt1@", Utils.AddComma(drs[i]["DoLosAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11EvatRslt1@", drs[i]["EvatRslt"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTbl평가결과, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTbl평가결과, oRow);
                        }
                        rUtil.TableMergeCellsV(oTbl평가결과, 0, rIdx1, rIdx2);


                        //교통비
                        drs = dtB?.Select("ExpsCd = 300263002 AND DmobSeq = " + ObjSeq); //교통비
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };
                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm2@");
                        oRowBase = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm2@");
                        rIdx1 = -1;
                        rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTbl평가결과, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTable(oTbl평가결과, "@B11DmobSortNo@", drs[i]["DmobSortNo"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ExpsGrpNm2@", drs[i]["ExpsGrpNm"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ReqAmt2@", Utils.AddComma(drs[i]["ReqAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11DoLosAmt2@", Utils.AddComma(drs[i]["DoLosAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11EvatRslt2@", drs[i]["EvatRslt"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTbl평가결과, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTbl평가결과, oRow);
                        }
                        rUtil.TableMergeCellsV(oTbl평가결과, 0, rIdx1, rIdx2);


                        //격락비
                        drs = dtB?.Select("ExpsCd = 300263003 AND DmobSeq = " + ObjSeq); //격락비
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };
                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm3@");
                        oRowBase = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm3@");
                        rIdx1 = -1;
                        rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTbl평가결과, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTable(oTbl평가결과, "@B11DmobSortNo@", drs[i]["DmobSortNo"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ExpsGrpNm3@", drs[i]["ExpsGrpNm"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ReqAmt3@", Utils.AddComma(drs[i]["ReqAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11DoLosAmt3@", Utils.AddComma(drs[i]["DoLosAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11EvatRslt3@", drs[i]["EvatRslt"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTbl평가결과, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTbl평가결과, oRow);
                        }
                        rUtil.TableMergeCellsV(oTbl평가결과, 0, rIdx1, rIdx2);


                        //기타
                        drs = dtB?.Select("ExpsCd = 300263004 AND DmobSeq = " + ObjSeq); //기타
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };
                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm4@");
                        oRowBase = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ExpsGrpNm4@");
                        rIdx1 = -1;
                        rIdx2 = -1;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTbl평가결과, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTable(oTbl평가결과, "@B11DmobSortNo@", drs[i]["DmobSortNo"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ExpsGrpNm4@", drs[i]["ExpsGrpNm"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B11ReqAmt4@", Utils.AddComma(drs[i]["ReqAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11DoLosAmt4@", Utils.AddComma(drs[i]["DoLosAmt"] + "") + "원");
                                rUtil.ReplaceTableRow(oRow, "@B11EvatRslt4@", drs[i]["EvatRslt"] + "");
                            }
                            if (i == 0) rIdx1 = rUtil.RowIndex(oTbl평가결과, oRow);
                            if (i == drs.Length - 1) rIdx2 = rUtil.RowIndex(oTbl평가결과, oRow);
                        }
                        rUtil.TableMergeCellsV(oTbl평가결과, 0, rIdx1, rIdx2);


                        //합계
                        drs = dtB?.Select("ExpsGrp = 91 AND DmobSeq = " + ObjSeq);
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };
                        
                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ReqAmt91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B11ReqAmt91@", Utils.AddComma(Utils.ToDouble(drs[0]["ReqAmt"] + "")) + "원");
                            rUtil.ReplaceTableRow(oRow, "@B11DoLosAmt91@", Utils.AddComma(Utils.ToDouble(drs[0]["DoLosAmt"] + "")) + "원");
                            rUtil.ReplaceTableRow(oRow, "@B11EvatRslt91@", drs[0]["EvatRslt"] + "");
                        }


                        //과실부담금
                        drs = dtB?.Select("ExpsGrp = 71 AND DmobSeq = " + ObjSeq);
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };

                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ReqAmt71@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B11ReqAmt71@", Utils.AddComma(Utils.ToDouble(drs[0]["ReqAmt"] + "")) + "원");
                            rUtil.ReplaceTableRow(oRow, "@B11EvatRslt71@", drs[0]["EvatRslt"] + "");
                        }


                        //자기부담금
                        drs = dtB?.Select("ExpsGrp = 75 AND DmobSeq = " + ObjSeq);
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };

                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ReqAmt75@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B11ReqAmt75@", Utils.AddComma(Utils.ToDouble(drs[0]["ReqAmt"] + "")) + "원");
                            rUtil.ReplaceTableRow(oRow, "@B11EvatRslt75@", drs[0]["EvatRslt"] + "");
                        }


                        //지급보험금
                        drs = dtB?.Select("ExpsGrp = 93 AND DmobSeq = " + ObjSeq);
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock11"].Rows.Add() };

                        oRow = rUtil.GetTableRow(oTbl평가결과?.Elements<TableRow>(), "@B11ReqAmt93@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B11ReqAmt93@", Utils.AddComma(Utils.ToDouble(drs[0]["ReqAmt"] + "")) + "원");
                            rUtil.ReplaceTableRow(oRow, "@B11EvatRslt93@", drs[0]["EvatRslt"] + "");
                        }
                    }


                    int db12RentAmtTot = 0;
                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    drs = dtB?.Select("DmobSeq = " + ObjSeq);
                    if (drs != null)
                    {
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock12"].Rows.Add() };
                        for (int i = 0; i < drs.Length; i++)
                        {
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = Utils.ConvertToString(drs[i][col] + "");
                                if (col.ColumnName == "FrDt") if (sValue != "") { sValue = Utils.DateFormat(sValue, "yyyy.MM.dd."); } //렌트 시작일
                                if (col.ColumnName == "ToDt") if (sValue != "") { sValue = "~ " + Utils.DateFormat(sValue, "yyyy.MM.dd."); } //렌트 종료일
                                if (col.ColumnName == "RentDay") if (sValue != "") { sValue = sValue + "일"; } //렌트일수
                                if (col.ColumnName == "RentTm") if (sValue != "") { sValue = sValue + "시간"; } //렌트시간
                                if (col.ColumnName == "RentAmt") if (Utils.ConvertToInt(sValue) > 0) { db12RentAmtTot += Utils.ConvertToInt(sValue); } //렌트비 합계
                                if (col.ColumnName == "RentAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; } //렌트비
                                rUtil.ReplaceTableRow(oTbl렌트비.GetRow(i + 2), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl렌트비, "@db12RentAmtTot@", Utils.AddComma(Utils.ConvertToString(db12RentAmtTot)) + "원");


                    int db13TrspAmtTot = 0;
                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    drs = dtB?.Select("DmobSeq = " + ObjSeq);
                    if (drs != null)
                    {
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        for (int i = 0; i < drs.Length; i++)
                        {
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = Utils.ConvertToString(drs[i][col] + "");
                                if (col.ColumnName == "FrDt") if (sValue != "") { sValue = Utils.DateFormat(sValue, "yyyy.MM.dd."); } //수리 시작일
                                if (col.ColumnName == "ToDt") if (sValue != "") { sValue = "~ " + Utils.DateFormat(sValue, "yyyy.MM.dd."); } //수리 종료일
                                if (col.ColumnName == "RentDay") if (sValue != "") { sValue = "( " + sValue + "일" + " )"; } //수리기간
                                if (col.ColumnName == "RentCost") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; } //1일 렌트요금
                                if (col.ColumnName == "TrspAmt") if (Utils.ConvertToInt(sValue) > 0) { db13TrspAmtTot += Utils.ConvertToInt(sValue); } //교통비 합계
                                if (col.ColumnName == "TrspAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; } //교통비 합계
                                TableRow oRow = rUtil.GetTableRow(oTbl교통비?.Elements<TableRow>(), sKey);
                                rUtil.ReplaceTableRow(oRow, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl교통비, "@db13TrspAmtTot@", Utils.AddComma(Utils.ConvertToString(db13TrspAmtTot)) + "원");
                    

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
