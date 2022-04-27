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
    public class RptAdjSLSurvRptGoodsKB_S_Building
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
                    Table oTbl보험가액 = rUtil.GetTable(lstTable, "db5EvatAmtTot1");
                    Table oTbl손해액 = rUtil.GetTable(lstTable, "db5EvatAmtTot2");
                    Table oTbl잔존물제거비용 = rUtil.GetTable(lstTable, "db5EvatAmtTot3");
                    


                    //보험가액
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 1");
                    if (drs != null)
                    {
                        if(oTbl보험가액 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl보험가액, 5, drs.Length - 1);
                            rUtil.TableMergeCells(oTbl보험가액, 0, 0, 5, drs.Length + 4);
                        }
                    }

                    //손해액
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 2");
                    if (drs != null)
                    {
                        if (oTbl손해액 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl손해액, 2, drs.Length - 1);
                            rUtil.TableMergeCells(oTbl손해액, 0, 0, 2, drs.Length + 1);
                        }
                    }

                    //잔존물제거비용
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 3");
                    if (drs != null)
                    {
                        if (oTbl잔존물제거비용 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl잔존물제거비용, 2, drs.Length - 1);
                            rUtil.TableMergeCells(oTbl잔존물제거비용, 0, 0, 2, drs.Length + 1);
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
                    Table oTbl보험가액 = rUtil.GetTable(lstTable, "db5EvatAmtTot1");
                    Table oTbl손해액 = rUtil.GetTable(lstTable, "db5EvatAmtTot2");
                    Table oTbl잔존물제거비용 = rUtil.GetTable(lstTable, "db5EvatAmtTot3");

                    ////변수가 replace 되기 전에 테이블을 찾아 놓는다
                    //sKey = rUtil.GetFieldName("B3", "EvatRsltRePurcTot");    //재조달가액 테이블
                    //Table oTableA = rUtil.GetSubTable(oTbl평가결과, sKey);
                    //sKey = rUtil.GetFieldName("B3", "ObjRstrTotal");         //복구공사비 테이블
                    //Table oTableB = rUtil.GetSubTable(oTbl평가결과, sKey);
                    //sKey = rUtil.GetFieldName("B3", "ObjRmnRmvTotal");       //잔존물제거비용 테이블
                    //Table oTableC = rUtil.GetSubTable(oTbl평가결과, sKey);

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        //총감가율
                        if (!dtB.Columns.Contains("TotDprcRate")) dtB.Columns.Add("TotDprcRate");
                        dr["TotDprcRate"] = Utils.Round((Utils.ToFloat(dr["EvatRsltPasDprcRate"]) / 12) * Utils.ToFloat(dr["EvatRsltPrgMm"]), 2);
                        ////재조달가액 합계
                        //if (!dtB.Columns.Contains("EvatRsltTotal")) dtB.Columns.Add("EvatRsltTotal");
                        //dr["EvatRsltTotal"] = Utils.ToFloat(dr["RePurcGexpAmt"]) + Utils.ToFloat(dr["EvatRsltRePurcTot"]);

                        ////복구공사비 합계
                        //if (!dtB.Columns.Contains("ObjRstrTotal")) dtB.Columns.Add("ObjRstrTotal");
                        //dr["ObjRstrTotal"] = Utils.ToFloat(dr["ObjRstrGexpTot"]) + Utils.ToFloat(dr["RstrGexpAmt"]);

                        ////잔존물제거비용 합계
                        //if (!dtB.Columns.Contains("ObjRmnRmvTotal")) dtB.Columns.Add("ObjRmnRmvTotal");
                        //dr["ObjRmnRmvTotal"] = Utils.ToFloat(dr["ObjRmnRmvTot"]) + Utils.ToFloat(dr["RmnObjRmvGexpAmt"]);

                        ////보험가액
                        //if (!dtB.Columns.Contains("EvatInsurTotal")) dtB.Columns.Add("EvatInsurTotal");
                        //if (!dtB.Columns.Contains("EvatRsltPrgYear")) dtB.Columns.Add("EvatRsltPrgYear");
                        //if (!dtB.Columns.Contains("EvatRsltPrgMonth")) dtB.Columns.Add("EvatRsltPrgMonth");
                        //double EvatRsltRePurcTot = Utils.ToDouble(dr["EvatRsltRePurcTot"]);
                        //double EvatRsltPasDprcRate = Utils.ToDouble(dr["EvatRsltPasDprcRate"]);
                        //double EvatRsltPrgMm = Utils.ToDouble(dr["EvatRsltPrgMm"]);
                        //double EvatRsltPrgYear = Math.Floor(EvatRsltPrgMm / 12);
                        //EvatRsltPrgMm = EvatRsltPrgMm % 12;
                        //double EvatInsurTotal = EvatRsltRePurcTot * (100.0 - (EvatRsltPasDprcRate * (EvatRsltPrgYear + EvatRsltPrgMm / 12))) / 100.0;
                        //dr["EvatInsurTotal"] = Utils.AddComma(EvatInsurTotal);
                        //dr["EvatRsltPrgYear"] = Utils.AddComma(EvatRsltPrgYear);
                        //dr["EvatRsltPrgMonth"] = Utils.AddComma(EvatRsltPrgMm);

                        ////손해액합계
                        //if (!dtB.Columns.Contains("DamageTotalAmt")) dtB.Columns.Add("DamageTotalAmt");
                        //double ObjRstrGexpTot = Utils.ToDouble(dr["ObjRstrGexpTot"]);
                        //double RstrGexpAmt = Utils.ToDouble(dr["RstrGexpAmt"]);
                        //double ObjRmnRmvTot = Utils.ToDouble(dr["ObjRmnRmvTot"]);
                        //double RmnObjRmvGexpAmt = Utils.ToDouble(dr["RmnObjRmvGexpAmt"]);
                        //dr["DamageTotalAmt"] = ObjRstrGexpTot + RstrGexpAmt + ObjRmnRmvTot + RmnObjRmvGexpAmt;

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "EvatRsltBuyDt") sValue = Utils.Mid(sValue, 1, 4) + "." + Utils.Mid(sValue, 5, 6); //건축시기
                            if (col.ColumnName == "EvatRsltPrgMm") //경과년수
                            {
                                sValue = Math.Floor(Utils.ConvertToDouble(sValue) / 12) + "년 " + (Utils.ConvertToDouble(sValue) % 12) + "월";
                            }
                            if (col.ColumnName == "EvatRsltTotArea") sValue = String.Format("{0:0.00}", Utils.ConvertToDouble(sValue)) + " ㎡"; //연면적
                            if (col.ColumnName == "EvatRsltPasDprcRate") sValue = sValue + "%"; //경년감가율
                            if (col.ColumnName == "TotDprcRate") sValue = sValue + "%"; //총감가율
                            ////재조달가액
                            //if (col.ColumnName == "RePurcGexpRate") sValue = sValue + "";
                            //if (col.ColumnName == "EvatRsltRePurcTot") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "RePurcGexpAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "EvatRsltTotal") sValue = Utils.AddComma(sValue);
                            ////복구공사비
                            //if (col.ColumnName == "RstrGexpRate") sValue = sValue + "";
                            //if (col.ColumnName == "ObjRstrGexpTot") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "RstrGexpAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "ObjRstrTotal") sValue = Utils.AddComma(sValue);
                            ////잔존물제거비용
                            //if (col.ColumnName == "RmnObjRmvGexpRate") sValue = sValue + "";
                            //if (col.ColumnName == "ObjRmnRmvTot") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "RmnObjRmvGexpAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "ObjRmnRmvTotal") sValue = Utils.AddComma(sValue);
                            ////보험가액
                            //if (col.ColumnName == "EvatInsurTotal") sValue = Utils.AddComma(sValue);
                            ////손해액합계
                            //if (col.ColumnName == "DamageTotalAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db5EvatAmtTot1 = 0;
                    double db5EvatAmtTot2 = 0;
                    double db5EvatAmtTot3 = 0;
                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        int ia = 0, ib = 0, ic = 0;
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            int EvatCd = Utils.ToInt(dtB.Rows[i]["EvatCd"]);

                            if (EvatCd % 10 == 1)  //보험가액
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "ObjInsValueTot") sValue = Utils.AddComma(sValue); //보험가액표 보험가액
                                    if (col.ColumnName == "EvatAmt") //보험가액표 합계
                                    {
                                        db5EvatAmtTot1 += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTbl보험가액.GetRow(ia + 5), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTbl보험가액.GetRow(ia + 7), sKey, sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl보험가액.GetRow(ia + 6), "@db5EvatAmtTot1@", Utils.AddComma(db5EvatAmtTot1)); //보험가액표 합계
                                
                                ia++;
                            }
                            else if (EvatCd % 10 == 2)  //손해액
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "EvatAmt") //손해액 합계
                                    {
                                        db5EvatAmtTot2 += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTbl손해액.GetRow(ib + 2), sKey, sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl손해액.GetRow(ib + 3), "@db5EvatAmtTot2@", Utils.AddComma(db5EvatAmtTot2)); //손해액표 합계
                                ib++;
                            }
                            else if (EvatCd % 10 == 3)  //잔존물제거비용
                            {
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "EvatAmt") //잔존물제거비용 합계
                                    {
                                        db5EvatAmtTot3 += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    rUtil.ReplaceTableRow(oTbl잔존물제거비용.GetRow(ic + 2), sKey, sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl잔존물제거비용.GetRow(ic + 3), "@db5EvatAmtTot3@", Utils.AddComma(db5EvatAmtTot3)); //잔존물제거비용표 합계
                                ic++;
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
