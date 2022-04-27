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
    public class RptAdjSLSurvRptGoodsNH_Head3
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                rUtil.TableAddRow(oTbl손해상황, 2, 1);
                            }
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@");

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        if (!dtB.Columns.Contains("AcdtDtTm")) dtB.Columns.Add("AcdtDtTm");
                        dr["AcdtDtTm"] = Utils.DateConv(dr["AcdtDt"] + "", ".") + " " + Utils.TimeConv(dr["AcdtTm"] + "", ":", "SHORT");
                        if (!dtB.Columns.Contains("AcdtJurdPolcText")) dtB.Columns.Add("AcdtJurdPolcText");
                        {
                            var AcdtJurdPolc = dr["AcdtJurdPolc"] + "";
                            var AcdtJurdPolcOpni = dr["AcdtJurdPolcOpni"] + "";
                            var AcdtJurdFire = dr["AcdtJurdFire"] + "";
                            var AcdtJurdFireOpni = dr["AcdtJurdFireOpni"] + "";
                            var text = AcdtJurdPolc + "\n" + AcdtJurdPolcOpni + "\n" + AcdtJurdFire + "\n" + AcdtJurdFireOpni;
                            dr["AcdtJurdPolcText"] = text;
                        }

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
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH시 mm분경");
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy. MM. dd.");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SealPhoto" || col.ColumnName == "ChrgAdjPhoto")
                            {
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.ReplaceInternalImage(sKey, img);
                                }
                                catch { }
                                continue;
                            }
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            int rcnt = 1;
                            string hdrBk = "";
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                                int rmdr = i % 2 + 1;

                                TableRow xrow1 = oTbl손해상황.GetRow(rnum);

                                sKey = rUtil.GetFieldName(sPrefix, "ObjNm");
                                sValue = dr["ObjNm"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);
                                //if (hdrBk != "" && hdrBk != sValue)
                                //{
                                //    rUtil.TableMergeCellsV(oTbl손해상황, 0, rcnt, rnum * 2 - 1);
                                //    rcnt = rnum;
                                //}
                                hdrBk = sValue;

                                sKey = rUtil.GetFieldName(sPrefix, "ObjSymb");
                                sValue = dr["ObjSymb"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2200000L, 1500000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictRmk");
                                sValue = dr["AcdtPictRmk"] + "";
                                rUtil.SetText(xrow1.GetCell(3), sKey, sValue);
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
