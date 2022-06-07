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
    public class RptAdjSLSurvMidRptGoodsNH_Head2
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
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl기계범례 = rUtil.GetTable(lstTable, "@B4ObjSymb_13@");

                    //3.일반사항 - 라.기계배치도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 3 OR ObjCatgCd % 10 = 4");
                    if (drs != null && drs.Length > 0)
                    {
                        //3.일반사항 - 라.기계배치도
                        if (oTbl기계범례 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl기계범례, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "B13AcdtPictImage");
                        if (oTbl기계배치도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl기계배치도, 0, 2, dtB.Rows.Count - 1);
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
                    Table oTbl기계배치도 = rUtil.GetTable(lstTable, "@B13AcdtPictImage@");
                    Table oTbl기계범례 = rUtil.GetTable(lstTable, "@B4ObjSymb_13@");

                    dtB = pds.Tables["DataBlock13"];
                    if (dtB == null || dtB.Rows.Count < 1)
                    {
                        //기계 배치도 사진 없으면 아래 기계범례 삭제
                        Table tbl = rUtil.GetTable(lstTable, "<기계범례>");
                        if (tbl != null) tbl.Remove();
                    }
                    else
                    {
                        //기계 배치도
                        drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 3 OR ObjCatgCd % 10 = 4");
                        sPrefix = "B4";
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                        if (drs != null && drs.Length > 0)
                        {
                            if (oTbl기계범례 != null)
                            {
                                for (int i = 0; i < drs.Length; i++)
                                {
                                    DataRow dr = drs[i];
                                    foreach (DataColumn col in dr.Table.Columns)
                                    {
                                        sKey = "@B4" + col.ColumnName + "_13@";
                                        sValue = dr[col] + "";
                                        if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                        if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                        if (col.ColumnName == "ObjBuyDt") sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년" + Utils.Mid(sValue, 5, 6) + "월");
                                        rUtil.ReplaceTableRow(oTbl기계범례.GetRow(i + 1), sKey, sValue);
                                    }
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        if (oTbl기계배치도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl기계배치도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl기계배치도.GetRow(rnum + 1);
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
