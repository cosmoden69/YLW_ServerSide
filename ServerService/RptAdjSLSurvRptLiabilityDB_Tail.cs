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
    public class RptAdjSLSurvRptLiabilityDB_Tail
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
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B13RmnObjCost@"); //잔존물 표1
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B13SucBidDt@"); //잔존물 표2
                    Table oTbl지급처 = rUtil.GetTable(lstTable, "@B5InsurGivObj@");
                    Table oTbl첨부자료목록 = rUtil.GetTable(lstTable, "@B4FileNo@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B7PrgMgtDt@");
                    

                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl당사, 1, drs.Length - 1);
                        }
                    }
                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }

                    //지급처
                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl지급처 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl지급처, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //첨부자료 목록
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (oTbl첨부자료목록 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl첨부자료목록, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //사고처리과정표 - 처리과정
                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl처리과정 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl처리과정, 1, dtB.Rows.Count - 1);
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
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B13RmnObjCost@"); //잔존물 표1
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B13SucBidDt@"); //잔존물 표2
                    Table oTbl지급처 = rUtil.GetTable(lstTable, "@B5InsurGivObj@");
                    Table oTbl첨부자료목록 = rUtil.GetTable(lstTable, "@B4FileNo@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B7PrgMgtDt@");
                    Table oTbl체크리스트 = rUtil.GetTable(lstTable, "@db12InsurObjDvs@");
                    Table oTbl기타사항 = rUtil.GetTable(lstTable, "@db8VitmText@");

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
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfpayAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
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

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (oTbl첨부자료목록 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "FileAmt") sValue = Utils.AddComma(sValue == "" ? "1" : sValue) + "부";
                                    rUtil.ReplaceTableRow(oTbl첨부자료목록.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl지급처 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl지급처.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    var db6OthInsur = "";
                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "OthInsurCo")
                                {
                                    var OthInsurCo = dr["OthInsurCo"] + "";
                                    var OthInsurPrdt = dr["OthInsurPrdt"] + "";
                                    db6OthInsur += OthInsurCo + ", " + OthInsurPrdt + "\n";
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db6OthInsur@", db6OthInsur);


                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
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

                    var db8VitmText = "";

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                
                                if (col.ColumnName == "VitmSubSeq")
                                {
                                    var VitmNm = dr["VitmNm"] + "";
                                    var VitmTel = dr["VitmTel"] + "";
                                    
                                    db8VitmText += VitmNm + " (" + VitmTel + ")" + "\n";
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl기타사항, "@db8VitmText@", db8VitmText); //사고처리과정표-기타사항

                    var db12InsurObjDvs = "";
                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //보험목적물
                                if (col.ColumnName == "ObjSymb")
                                {
                                    var InsurObjDvs = dr["InsurObjDvs"] + "";
                                    db12InsurObjDvs += InsurObjDvs + "\n";
                                }

                                //담보여부
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "1")
                                    {
                                        sValue = "보험 가입";
                                    }
                                    else
                                    {
                                        sValue = "보험 미가입";
                                    }
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db12InsurObjDvs@", db12InsurObjDvs);


                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 1");
                    sPrefix = "B13";
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
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
                                    rUtil.ReplaceTableRow(oTbl당사.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl당사.Remove();
                    }

                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 2");
                    sPrefix = "B13";
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
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "AuctFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "AuctToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "SucBidDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTbl옥션.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl옥션.Remove();
                    }

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
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

                    var db14InsurObjNm = "";
                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //미가입목적물
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "0" || sValue == "")
                                    {
                                        var InsurObjNm = dr["InsurObjNm"] + "";
                                        db14InsurObjNm += InsurObjNm + "\n";
                                    }

                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTables(lstTable, sKey, sValue);
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db14InsurObjNm@", db14InsurObjNm);

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
