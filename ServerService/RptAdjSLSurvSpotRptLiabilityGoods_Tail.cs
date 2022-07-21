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
    public class RptAdjSLSurvSpotRptLiabilityGoods_Tail
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
                    Table oTbl잔존물 = rUtil.GetTable(lstTable, "@B16RmnObjNm@");
                    Table oTbl타보험 = rUtil.GetTable(lstTable, "@B5OthInsurCo@");

                    dtB = pds.Tables["DataBlock16"];
                    if (dtB != null)
                    {
                        if (oTbl잔존물 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl잔존물, 1, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    if (dtB != null)
                    {
                        if (oTbl타보험 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl타보험, 2, 2, dtB.Rows.Count - 1);
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
                    Table oTbl잔존물 = rUtil.GetTable(lstTable, "@B16RmnObjNm@");
                    Table oTbl가치없음 = rUtil.GetTable(lstTable, "경제적 잔존가치 없음.");
                    Table oTbl타보험 = rUtil.GetTable(lstTable, "@B5OthInsurCo@");
                    Table oTbl해당없음 = rUtil.GetTable(lstTable, "해당사항 없음.");

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
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock16"];
                    sPrefix = "B16";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        oTbl가치없음.Remove();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "AuctFrDt")
                                {
                                    sValue = Utils.DateFormat(dr["AuctFrDt"], "yyyy.MM.dd") + "\n ~" + Utils.DateFormat(dr["AuctToDt"], "yyyy.MM.dd");
                                    sKey = "@db16AuctTerms@";
                                }
                                if (col.ColumnName == "AuctToDt") continue;
                                if (col.ColumnName == "SucBidDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                TableRow oRow = rUtil.GetTableRow(oTbl잔존물?.Elements<TableRow>(), sKey);
                                rUtil.ReplaceTableRow(oRow, sKey, sValue);
                            }
                        }
                    }
                    else
                    {
                        oTbl잔존물.Remove();
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        oTbl해당없음.Remove();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            int rnum = (i + 1) * 2;
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd") + " ~";
                                if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "OthSelfBearAmt") sValue = Utils.AddComma(sValue);
                                rUtil.ReplaceTableRow(oTbl타보험.GetRow(rnum + 0), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl타보험.GetRow(rnum + 1), sKey, sValue);
                            }
                        }
                    }
                    else
                    {
                        oTbl타보험.Remove();
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
