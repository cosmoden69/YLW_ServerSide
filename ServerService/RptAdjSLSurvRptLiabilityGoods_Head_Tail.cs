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
    public class RptAdjSLSurvRptLiabilityGoods_Head_Tail
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
                    Table oTbl타보험 = rUtil.GetTable(lstTable, "@B5OthInsurCo@");
                    Table oTbl일반사항 = rUtil.GetTable(lstTable, "@B6Insured@");

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
                    Table oTbl타보험 = rUtil.GetTable(lstTable, "@B5OthInsurCo@");
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
                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
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
