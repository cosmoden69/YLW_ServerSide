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
    public class RptAdjSLSurvDmgRptGoods_Tail
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
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");

                    //잔존물가액 2
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs != null || drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "PrgMgtDt");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPrsCcndGrp");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTable, 1, dtB.Rows.Count - 1);
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
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");

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
                                if (sValue == "") sValue = "-";
                                if (dr["RgtCpstOpni"] + "" != "") sValue += "\n" + dr["RgtCpstOpni"];
                            }
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
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
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
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
                                    rUtil.ReplaceTableRow(oTbl옥션.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
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
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "PrgMgtDt");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateConv(sValue, ".");
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPrsCcndGrp");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
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
