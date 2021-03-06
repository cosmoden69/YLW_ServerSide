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
    public class RptAdjSLSurvRptGoodsNH_Tail
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
                    Table oTbl잔존물_표1 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl잔존물_표2 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl관련자연락처 = rUtil.GetTable(lstTable, "@B11AcdtPrsCcndGrp@");
                    Table oTbl유첨 = rUtil.GetTable(lstTable, "@B9FileNo@");
                    Table oTbl첨부사진 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");

                    //1.잔존물_표1
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    if (drs != null || drs.Length > 0)
                    {
                        if (oTbl잔존물_표1 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl잔존물_표1, 1, drs.Length - 1);
                        }
                    }

                    //1.잔존물_표2
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs != null || drs.Length > 0)
                    {
                        if (oTbl잔존물_표2 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl잔존물_표2, 1, drs.Length - 1);
                        }
                    }

                    //유첨
                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl유첨 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl유첨, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //사고처리과정표 - 처리과정
                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl처리과정 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl처리과정, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //2-3.관련자연락처
                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (oTbl관련자연락처 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl관련자연락처, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl첨부사진 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl첨부사진, 0, 2, dtB.Rows.Count - 1);
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
                    Table oTbl잔존물_표1 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl잔존물_표2 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl관련자연락처 = rUtil.GetTable(lstTable, "@B11AcdtPrsCcndGrp@");
                    Table oTbl유첨 = rUtil.GetTable(lstTable, "@B9FileNo@");
                    Table oTbl첨부사진 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    sKey = rUtil.GetFieldName("B8", "RmnObjStd");
                    Table oTableA = rUtil.GetTable(lstTable, sKey);

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
                            if (col.ColumnName == "RgtCpstCnclsRmk")
                            {
                                if (dr["RgtCpstOpni"] + "" != "") sValue += (sValue != "" ? "\n" : "") + dr["RgtCpstOpni"];
                                if (sValue == "") sValue = "-";
                            }
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH시 mm분");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    sPrefix = "B8";
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl잔존물_표1 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl잔존물_표1.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl잔존물_표2 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "AuctFrDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "AuctToDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "SucBidDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl잔존물_표2.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl유첨 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "FileAmt") sValue = Utils.AddComma(sValue == "" || sValue == "0" ? "1" : sValue) + "부";
                                    rUtil.ReplaceTableRow(oTbl유첨.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl처리과정 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                                    rUtil.ReplaceTableRow(oTbl처리과정.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (oTbl관련자연락처 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTbl관련자연락처.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }


                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        DataTable dtX = dtB.Copy();
                        if (oTbl첨부사진 != null)
                        {
                            if (dtX.Rows.Count < 1) dtX.Rows.Add();
                            for (int i = 0; i < dtX.Rows.Count; i++)
                            {
                                DataRow dr = dtX.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl첨부사진.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 4200000L, 3300000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl첨부사진.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
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
