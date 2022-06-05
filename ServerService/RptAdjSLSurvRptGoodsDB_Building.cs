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
    class RptAdjSLSurvRptGoodsDB_Building
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
                    Table oTbl감가상각 = rUtil.GetTable(lstTable, "@B3InsurObjDvs@");
                    Table oTbl신축비_수리비 = rUtil.GetTable(lstTable, "@B3RePurcGexpAmt@");
                    Table oTbl잔존물제거비용A = rUtil.GetTable(lstTable, "@B0RemainsA@");
                    Table oTbl잔존물제거비 = rUtil.GetSubTable(oTbl잔존물제거비용A, "@B3RmnObjRmvGexpAmt@");
                    Table oTbl잔존물제거비용B = rUtil.GetTable(lstTable, "@B0RemainsB@");

                    //1-3.감가상각
                    dtB = pds.Tables["DataBlock3"];
                    if (dtB != null)
                    {
                        if (oTbl감가상각 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl감가상각, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //2-8.신축비/수리비
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 1");
                    if (drs == null || drs.Length < 1)
                    {
                        if (oTbl신축비_수리비 != null) rUtil.TableRemoveRow(oTbl신축비_수리비, 1);
                    }
                    else
                    {
                        if (oTbl신축비_수리비 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl신축비_수리비, 1, drs.Length - 1);
                        }
                    }
                    

                    //2-9.잔존물제거비
                    drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 3 or EvatCd % 10 = 4 or EvatCd % 10 = 5");
                    if (drs == null || drs.Length < 1)
                    {
                        oTbl잔존물제거비용A.Remove();
                    }
                    else
                    {
                        oTbl잔존물제거비용B?.Remove();
                        if (oTbl잔존물제거비 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl잔존물제거비, 1, drs.Length - 1);
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
                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    Table oTbl감가상각 = rUtil.GetTable(lstTable, "@B3InsurObjDvs@");
                    Table oTbl신축비_수리비 = rUtil.GetTable(lstTable, "@B3RePurcGexpAmt@");
                    Table oTbl잔존물제거비용A = rUtil.GetTable(lstTable, "@B0RemainsA@");
                    Table oTbl잔존물제거비 = rUtil.GetSubTable(oTbl잔존물제거비용A, "@B3RmnObjRmvGexpAmt@");
                    Table oTbl잔존물제거비용B = rUtil.GetTable(lstTable, "@B0RemainsB@");


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
                        //신축비/수리비
                        if (!dtB.Columns.Contains("EvatRsltTotal")) dtB.Columns.Add("EvatRsltTotal");
                        dr["EvatRsltTotal"] = dr["EvatRsltRePurcTot"];  // + Utils.ToFloat(dr["RePurcGexpAmt"]);

                        //잔존물제거비용 합계
                        if (!dtB.Columns.Contains("ObjRmnRmvTotal")) dtB.Columns.Add("ObjRmnRmvTotal");
                        dr["ObjRmnRmvTotal"] = dr["ObjRmnRmvTot"];   // + Utils.ToDecimal(dr["RmnObjRmvGexpAmt"]);

                        //손해액합계
                        if (!dtB.Columns.Contains("DmgAmtTot")) dtB.Columns.Add("DmgAmtTot");
                        double ObjLosAmt = Utils.ToDouble(dr["ObjLosAmt"]);
                        double ObjRmnRmvTot = Utils.ToDouble(dr["ObjRmnRmvTot"]);
                        double RmnObjRmvGexpAmt = Utils.ToDouble(dr["RmnObjRmvGexpAmt"]);
                        dr["DmgAmtTot"] = ObjLosAmt + ObjRmnRmvTot + RmnObjRmvGexpAmt;

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjSymb")
                            {
                                while (Utils.Left(sValue, 1) == ",")
                                {
                                    sValue = Utils.Mid(sValue, 2, sValue.Length);
                                }
                            }
                            if (col.ColumnName == "EvatRsltBuyDt") sValue = Utils.Mid(sValue, 1, 4) + "." + Utils.Mid(sValue, 5, 6);
                            if (col.ColumnName == "EvatRsltPrgMm")
                            {
                                sValue = Math.Floor(Utils.ConvertToDouble(sValue) / 12) + "년 " + (Utils.ConvertToDouble(sValue) % 12) + "월";
                            }
                            if (col.ColumnName == "EvatRsltTotArea") sValue = String.Format("{0:0.00}", Utils.ConvertToDouble(sValue)) + " ㎡";
                            if (col.ColumnName == "EvatRsltPasDprcRate") sValue = sValue + "%";
                            if (col.ColumnName == "ObjInsValueTot") sValue = Utils.AddComma(sValue);//보험가액
                            if (col.ColumnName == "ObjLosAmt") sValue = Utils.AddComma(sValue);//손해액
                            //신축비/수리비
                            if (col.ColumnName == "RePurcGexpRate") sValue = sValue + "";
                            if (col.ColumnName == "EvatRsltRePurcTot") continue;
                            if (col.ColumnName == "RePurcGexpAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "EvatRsltTotal") sValue = Utils.AddComma(sValue);
                            //잔존물제거비용
                            if (col.ColumnName == "RmnObjRmvGexpRate") sValue = sValue + "";
                            if (col.ColumnName == "ObjRmnRmvTot") continue;
                            if (col.ColumnName == "RmnObjRmvGexpAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjRmnRmvTotal") sValue = Utils.AddComma(sValue);
                            //손해액합계
                            if (col.ColumnName == "DmgAmtTot") sValue = Utils.AddComma(sValue);
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
                        //잔존물제거비용 합계
                        double ObjRmnRmvTot = 0;
                        int ia = 0, ib = 0, ic = 0;
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            int EvatCd = Utils.ToInt(dtB.Rows[i]["EvatCd"]);

                            if (EvatCd % 10 == 1)  //신축비/수리비
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
                                    rUtil.ReplaceTableRow(oTbl신축비_수리비.GetRow(ia + 1), sKey, sValue);
                                }
                                ia++;
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
                                    rUtil.ReplaceTableRow(oTbl잔존물제거비.GetRow(ic + 1), sKey, sValue);
                                }
                                ic++;
                            }
                        }
                        rUtil.ReplaceTable(oTbl신축비_수리비, "@B3EvatRsltRePurcTot@", Utils.AddComma(EvatRsltRePurcTot));
                        rUtil.ReplaceTable(oTbl잔존물제거비, "@B3ObjRmnRmvTot@", Utils.AddComma(ObjRmnRmvTot));
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
