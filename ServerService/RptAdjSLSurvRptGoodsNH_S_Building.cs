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
    public class RptAdjSLSurvRptGoodsNH_S_Building
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
                    Table oTbl신축가액 = rUtil.GetTable(lstTable, "신축가액");
                    Table oTbl손해액 = rUtil.GetTable(lstTable, "@B3ObjRstrTotal@");
                    Table oTbl잔존물제거비용 = rUtil.GetTable(lstTable, "@B3ObjRmnRmvTotal@");
                    Table oTbl잔존물제거비용B = rUtil.GetTable(lstTable, "@B0RemainsB@");


                    // 신축가액
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 1");
                    sKey = rUtil.GetFieldName("B3", "EvatRsltTotal");
                    Table oTableA = rUtil.GetSubTable(oTbl신축가액, sKey);
                    if (drs == null || drs.Length < 1)
                    {
                        if (oTableA != null) rUtil.TableRemoveRow(oTableA, 1);
                    }
                    else
                    {
                        if (oTableA != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTableA, 1, drs.Length - 1);
                        }
                    }

                    //손해액 행추가
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 2");
                    if (drs == null || drs.Length < 1)
                    {
                        if (oTbl손해액 != null) rUtil.TableRemoveRow(oTbl손해액, 1);
                    }
                    else
                    {
                        if (oTbl손해액 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl손해액, 1, drs.Length - 1);
                        }
                    }

                    //잔존물제거비 행추가
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 3");
                    if (drs == null || drs.Length < 1)
                    {
                        oTbl잔존물제거비용.Remove();
                    }
                    else
                    {
                        oTbl잔존물제거비용B?.Remove();
                        if (oTbl잔존물제거비용 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl잔존물제거비용, 1, drs.Length - 1);
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
                    Table oTbl신축가액 = rUtil.GetTable(lstTable, "신축가액");
                    Table oTbl손해액 = rUtil.GetTable(lstTable, "@B3ObjRstrTotal@");
                    Table oTbl잔존물제거비용 = rUtil.GetTable(lstTable, "@B3ObjRmnRmvTotal@");
                    Table oTbl총감가율 = rUtil.GetTable(lstTable, "@db3EvatRsltPasDprcRate@");
                    Table oTbl잔존물제거비용B = rUtil.GetTable(lstTable, "@B0RemainsB@");

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    Table oTableA = rUtil.GetSubTable(oTbl신축가액, "@B3EvatRsltTotal@");

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
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
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

                        //신축가액
                        if (!dtB.Columns.Contains("EvatRsltTotal")) dtB.Columns.Add("EvatRsltTotal");
                        dr["EvatRsltTotal"] = dr["EvatRsltRePurcTot"];  // + Utils.ToFloat(dr["RePurcGexpAmt"]);

                        //손해액(수리비) 합계
                        if (!dtB.Columns.Contains("ObjRstrTotal")) dtB.Columns.Add("ObjRstrTotal");
                        dr["ObjRstrTotal"] = dr["ObjRstrGexpTot"];  // + Utils.ToFloat(dr["RstrGexpAmt"]);

                        //잔존물제거비용 합계
                        if (!dtB.Columns.Contains("ObjRmnRmvTotal")) dtB.Columns.Add("ObjRmnRmvTotal");
                        dr["ObjRmnRmvTotal"] = dr["ObjRmnRmvTot"];  // + Utils.ToDecimal(dr["RmnObjRmvGexpAmt"]);

                        //보험가액
                        if (!dtB.Columns.Contains("EvatInsurTotal")) dtB.Columns.Add("EvatInsurTotal");
                        if (!dtB.Columns.Contains("EvatRsltPrgYear")) dtB.Columns.Add("EvatRsltPrgYear");
                        if (!dtB.Columns.Contains("EvatRsltPrgMonth")) dtB.Columns.Add("EvatRsltPrgMonth");
                        double EvatRsltTotal = Utils.ToDouble(dr["EvatRsltTotal"]);
                        double EvatRsltPasDprcRate = Utils.ToDouble(dr["EvatRsltPasDprcRate"]);
                        double EvatRsltPrgMm = Utils.ToDouble(dr["EvatRsltPrgMm"]);
                        double EvatRsltPrgYear = Math.Floor(EvatRsltPrgMm / 12);
                        EvatRsltPrgMm = EvatRsltPrgMm % 12;
                        double EvatInsurTotal = EvatRsltTotal * (100.0 - (EvatRsltPasDprcRate * (EvatRsltPrgYear + EvatRsltPrgMm / 12))) / 100.0;
                        dr["EvatInsurTotal"] = Utils.AddComma(EvatInsurTotal);
                        dr["EvatRsltPrgYear"] = Utils.AddComma(EvatRsltPrgYear);
                        dr["EvatRsltPrgMonth"] = Utils.AddComma(EvatRsltPrgMm);

                        //감가상각
                        if (!dtB.Columns.Contains("RstrTotalAmt")) dtB.Columns.Add("RstrTotalAmt");
                        double ObjRstrTotal = Utils.ToDouble(dr["ObjRstrTotal"]);
                        double RstrTotalAmt = ObjRstrTotal * (100.0 - (EvatRsltPasDprcRate * (EvatRsltPrgYear + EvatRsltPrgMm / 12))) / 100.0;
                        dr["RstrTotalAmt"] = Utils.AddComma(RstrTotalAmt);

                        //손해액합계
                        if (!dtB.Columns.Contains("RstrTotalAmt10")) dtB.Columns.Add("RstrTotalAmt10");
                        if (!dtB.Columns.Contains("DamageTotalAmt")) dtB.Columns.Add("DamageTotalAmt");
                        double RstrTotalAmt10 = Math.Round(RstrTotalAmt * 0.1);
                        double DamageTotalAmt = RstrTotalAmt + RstrTotalAmt10;
                        dr["RstrTotalAmt10"] = RstrTotalAmt10;
                        dr["DamageTotalAmt"] = DamageTotalAmt;

                        //총감가율
                        if (!dtB.Columns.Contains("TotEvatRsltPasDprcRate")) dtB.Columns.Add("TotEvatRsltPasDprcRate");
                        {
                            dr["TotEvatRsltPasDprcRate"] = Utils.Round(EvatRsltPasDprcRate * (EvatRsltPrgYear + EvatRsltPrgMm / 12), 2);
                        }

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "EvatRsltBuyDt") sValue = Utils.Mid(sValue, 1, 4) + "." + Utils.Mid(sValue, 5, 6);
                            if (col.ColumnName == "EvatRsltPrgMm")
                            {
                                sValue = Math.Floor(Utils.ConvertToDouble(sValue) / 12) + "년 " + (Utils.ConvertToDouble(sValue) % 12) + "월";
                            }
                            if (col.ColumnName == "EvatRsltTotArea") sValue = String.Format("{0:0.00}", Utils.ConvertToDouble(sValue)) + " ㎡";
                            if (col.ColumnName == "EvatRsltPasDprcRate") sValue = sValue + "%";
                            //신축가액
                            if (col.ColumnName == "RePurcGexpRate") sValue = sValue + "";
                            if (col.ColumnName == "EvatRsltRePurcTot") continue;
                            if (col.ColumnName == "RePurcGexpAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "EvatRsltTotal") sValue = Utils.AddComma(sValue);
                            //손해액(수리비)
                            if (col.ColumnName == "RstrGexpRate") sValue = sValue + "";
                            if (col.ColumnName == "ObjRstrGexpTot") continue;
                            if (col.ColumnName == "RstrGexpAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjRstrTotal") sValue = Utils.AddComma(sValue);
                            //잔존물제거비용
                            if (col.ColumnName == "RmnObjRmvGexpRate") sValue = sValue + "";
                            if (col.ColumnName == "ObjRmnRmvTot") continue;
                            if (col.ColumnName == "RmnObjRmvGexpAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjRmnRmvTotal") sValue = Utils.AddComma(sValue);
                            //보험가액
                            if (col.ColumnName == "EvatInsurTotal") sValue = Utils.AddComma(sValue);
                            //손해액합계
                            if (col.ColumnName == "ObjLosAmt") sValue = Utils.AddComma(sValue);
                            ////총감가율
                            if (col.ColumnName == "TotEvatRsltPasDprcRate") sValue = sValue + "%";
                            //손해액합계
                            if (col.ColumnName == "DamageTotalAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        //신축비/수리비 소계
                        double EvatRsltRePurcTot = 0;
                        //복구공사비 소계
                        double ObjRstrGexpTot = 0;
                        //잔존물제거비용 합계
                        double ObjRmnRmvTot = 0;
                        int ia = 0, ib = 0, ic = 0;
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            int EvatCd = Utils.ToInt(dtB.Rows[i]["EvatCd"]);

                            if (EvatCd % 10 == 1)  //신축가액
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt")
                                    {
                                        sValue = Utils.AddComma(sValue);
                                        EvatRsltRePurcTot += Utils.ToDouble(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTableA.GetRow(ia + 1), sKey, sValue);
                                }
                                ia++;
                            }
                            if (EvatCd % 10 == 2)  //손해액
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt")
                                    {
                                        sValue = Utils.AddComma(sValue);
                                        ObjRstrGexpTot += Utils.ToDouble(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTbl손해액.GetRow(ib + 1), sKey, sValue);
                                }
                                ib++;
                            }
                            else if (EvatCd % 10 == 3)  //잔존물제거비용
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt")
                                    {
                                        sValue = Utils.AddComma(sValue);
                                        ObjRmnRmvTot += Utils.ToDouble(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTbl잔존물제거비용.GetRow(ic + 1), sKey, sValue);
                                }
                                ic++;
                            }
                        }
                        if (ia > 0) rUtil.TableMergeCells(oTableA, 0, 0, 1, ia);

                        rUtil.ReplaceTable(oTableA, "@B3EvatRsltRePurcTot@", Utils.AddComma(EvatRsltRePurcTot));
                        rUtil.ReplaceTable(oTbl손해액, "@B3ObjRstrGexpTot@", Utils.AddComma(ObjRstrGexpTot));
                        rUtil.ReplaceTable(oTbl잔존물제거비용, "@B3ObjRmnRmvTot@", Utils.AddComma(ObjRmnRmvTot));
                    }

                    rUtil.ReplaceTables(lstTable, "@B0RemainsA@", "");
                    rUtil.ReplaceTables(lstTable, "@B0RemainsB@", "");

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
