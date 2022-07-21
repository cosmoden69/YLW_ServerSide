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
    public class RptAdjSLSurvSpotRptLiabilityGoods_Head
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B4InsurPrdt@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtSurvDsnt@");
                    Table oTbl일반사항 = rUtil.GetTable(lstTable, "@B6Insured@");

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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B4InsurPrdt@");
                    Table oTbl사고내용 = rUtil.GetTable(lstTable, "@B1AcdtSurvDsnt@");
                    Table oTbl일반사항 = rUtil.GetTable(lstTable, "@B6Insured@");

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
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            TableRow oRow = rUtil.GetTableRow(oTbl사고내용?.Elements<TableRow>(), sKey);
                            rUtil.ReplaceTableRow(oRow, sKey, sValue);
                            if (col.ColumnName == "AcdtJurdFire" && sValue.Trim() == "") rUtil.TableRemoveRow(oTbl사고내용, oRow);
                            if (col.ColumnName == "AcdtJurdPolc" && sValue.Trim() == "") rUtil.TableRemoveRow(oTbl사고내용, oRow);
                            if (col.ColumnName == "SurvOpni" && sValue.Trim() == "") rUtil.TableRemoveRow(oTbl사고내용, oRow);
                            if (col.ColumnName == "AcdtSurvDsnt" && sValue.Trim() == "") rUtil.TableRemoveRow(oTbl사고내용, oRow);
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    string db4InsurPrdt = "";
                    string db4InsurNo = "";
                    string db4CtrtDt = "";
                    string db4CltrSpcCtrt = "";
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
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
                                if (col.ColumnName == "InsurPrdt")
                                {
                                    if (i > 0) db4InsurPrdt += "\n";
                                    db4InsurPrdt += sValue;
                                }
                                if (col.ColumnName == "InsurNo")
                                {
                                    if (i > 0) db4InsurNo += "\n";
                                    db4InsurNo += sValue;
                                }
                                if (col.ColumnName == "CtrtDt")
                                {
                                    if (i > 0) db4CtrtDt += "\n";
                                    db4CtrtDt += Utils.DateFormat(dr["CtrtDt"], "yyyy.MM.dd") + " ~ " + Utils.DateFormat(dr["CtrtExprDt"], "yyyy.MM.dd");
                                }
                                if (col.ColumnName == "CltrSpcCtrt")
                                {
                                    if (i > 0) db4CltrSpcCtrt += "\n";
                                    db4CltrSpcCtrt += sValue;
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl계약사항, "@B4InsurPrdt@", db4InsurPrdt);
                    rUtil.ReplaceTable(oTbl계약사항, "@B4InsurNo@", db4InsurNo);
                    rUtil.ReplaceTable(oTbl계약사항, "@B4CtrtDt@", db4CtrtDt);
                    rUtil.ReplaceTable(oTbl계약사항, "@B4CltrSpcCtrt@", db4CltrSpcCtrt);

                    string db7VitmNm = "";
                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
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
                                if (col.ColumnName == "Vitm")
                                {
                                    if (i > 0) db7VitmNm += "\n";
                                    db7VitmNm += sValue;
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl계약사항, "@B7VitmNm@", db7VitmNm);

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
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
                                if (col.ColumnName == "IsrdOpenDt") sValue = Utils.SubString(sValue, 0, 4) + "." + Utils.SubString(sValue, 4, 2);
                                if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue) + "원";
                                TableRow oRow = rUtil.GetTableRow(oTbl일반사항?.Elements<TableRow>(), sKey);
                                rUtil.ReplaceTableRow(oRow, sKey, sValue);
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
