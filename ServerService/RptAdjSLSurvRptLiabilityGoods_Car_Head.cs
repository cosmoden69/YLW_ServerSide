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
    public class RptAdjSLSurvRptLiabilityGoods_Car_Head
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


                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "DmobSortNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }


                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "InsurGivObj");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }
                    
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "OthInsurCo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTable, 2, 2, dtB.Rows.Count - 1);

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
                    Table oTable총괄표 = rUtil.GetTable(lstTable, "@B2DmobSortNo@");
                    Table oTable보험금지급처 = rUtil.GetTable(lstTable, "@B3InsurGivObj@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B4OthInsurCo@");
                    
                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (!dtB.Columns.Contains("DoOthExpsHedText")) dtB.Columns.Add("DoOthExpsHedText");
                        {
                            if (Utils.ConvertToString(dr["DoOthExpsHed"]) == "")
                            {
                                dr["DoOthExpsHedText"] = "4. ";
                            }
                            else
                            {
                                dr["DoOthExpsHedText"] = "4." + dr["DoOthExpsHed"];
                            }
                        }

                        if (!dtB.Columns.Contains("DoOthExpsHedText")) dtB.Columns.Add("DoOthExpsHedText");
                        {
                            if ((Utils.ConvertToInt(dr["DoOthExpsReq"]) == 0) && (Utils.ConvertToString(dr["DoOthExpsReq"]) == "") && (Utils.ConvertToInt(dr["DoOthExpsAmt"]) == 0) && (Utils.ConvertToString(dr["DoOthExpsAmt"]) == ""))
                            {
                                dr["DoOthExpsHedText"] = " ";
                                dr["DoOthExpsReq"] = 0;
                                dr["DoOthExpsAmt"] = 0;
                                dr["DoOthExpsCmnt"] = " ";
                                dr["DoOthExpsBss"] = " ";
                            }
                        }

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "DeptName") sValue = (sValue == "" ? "-" : sValue);
                            if (col.ColumnName == "EmpWorkAddress") sValue = (sValue == "" ? "-" : sValue);
                            if (col.ColumnName == "DeptPhone") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "DeptFax") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "IsrtTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "IsrdTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd") + " ~";
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "FixFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FixToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            //if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "InsurRegsAmt2") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotReq") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AgrmAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoBivInsurAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue); 삭제
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InsurRegsAmtRevw") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmtRevw") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "DoFixReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoFixAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSubTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurAmt") sValue = Utils.AddComma(sValue);
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

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }


                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
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
                                if (col.ColumnName == "CarTyp") if (sValue != "") { sValue = "(" + sValue + ")"; } //차량종류
                                if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue) + "원"; //보상한도액
                                if (col.ColumnName == "ReqAmt") sValue = Utils.AddComma(sValue) + "원"; //청구액
                                if (col.ColumnName == "DoLosAmt") sValue = Utils.AddComma(sValue) + "원"; //손해액
                                if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue) + "원"; //자기부담금
                                if (col.ColumnName == "GivInsurAmt") sValue = Utils.AddComma(sValue) + "원"; //지급보험금  
                                rUtil.ReplaceTableRow(oTable총괄표.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }


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
                                if (col.ColumnName == "GivObjRegno") if (sValue != "") { sValue = "(" + sValue + ")"; } //주민번호(사업자번호)
                                if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue) + "원   "; //지급보험금
                                rUtil.ReplaceTableRow(oTable보험금지급처.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }


                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
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
                                    if (col.ColumnName == "OthInsurRegsAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; }
                                    if (col.ColumnName == "OthSelfBearAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; }
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 1), sKey, sValue);
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
