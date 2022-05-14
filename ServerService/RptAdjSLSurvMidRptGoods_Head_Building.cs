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
    public class RptAdjSLSurvMidRptGoods_Head_Building
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
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");

                    //건물구조 및 면적
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        //3.일반사항 - 다.건물현황 및 배치도
                        Table oTableE = rUtil.GetSubTable(oTbl건물현황및배치도, "@B4ObjSymb_12@");
                        if (oTableE != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTableE, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        //3.일반사항 - 다.건물현황 및 배치도
                        Table oTableG = rUtil.GetSubTable(oTbl건물현황및배치도, "@B7AcdtPictImage@");
                        if (oTableG != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRows(oTableG, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    if (dtB != null)
                    {
                        Table oTableH = rUtil.GetSubTable(oTbl건물현황및배치도, "@B12AcdtPictImage@");
                        if (oTableH != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 2) / 3.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableH, 0, 1);
                                rUtil.TableAddRow(oTableH, 1, 1);
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
                    Table oTbl건물현황및배치도 = rUtil.GetTable(lstTable, "@B7AcdtPictImage@");

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    Table oTableE = rUtil.GetSubTable(oTbl건물현황및배치도, "@B4ObjSymb_12@");      //3.일반사항 - 다.건물현황 및 배치도
                    Table oTableG = rUtil.GetSubTable(oTbl건물현황및배치도, "@B7AcdtPictImage@");   //3.일반사항 - 다.건물현황 및 배치도
                    Table oTableH = rUtil.GetSubTable(oTbl건물현황및배치도, "@B12AcdtPictImage@");

                    //건물현황 및 배치도
                    drs = pds.Tables["DataBlock4"]?.Select("ObjCatgCd % 10 = 1");
                    sPrefix = "B4";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock4"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTableE != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = "@B4" + col.ColumnName + "_12@";
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjSymb") sValue = sValue.Replace(",", "");
                                    if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTableE.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTableG != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableG.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableG.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (oTableH != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 3 != 0)
                            {
                                if (dtB.Rows.Count % 3 >= 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                                if (dtB.Rows.Count % 3 >= 2) dtB.Rows.Add();  //세번째 칸을 클리어 해주기 위해서 추가
                            }
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 3.0) * 2;
                                int rmdr = i % 3;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableH.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2000000L, 1400000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableH.GetRow(rnum + 1);
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
