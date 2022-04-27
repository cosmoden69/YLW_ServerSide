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
    public class RptAdjSLSurvRptGoodsNH_Object
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

                    //평가결과 행추가
                    drs = pds.Tables["DataBlock4"]?.Select("1 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        sKey = rUtil.GetFieldName("B4", "RprtSeq");
                        Table oTableA = rUtil.GetTable(lstTable, sKey);
                        if (oTableA != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTableA, 2, 2, drs.Length - 1);
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
                    sKey = rUtil.GetFieldName("B4", "RprtSeq");
                    Table oTableA = rUtil.GetTable(lstTable, sKey);

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
                            if (col.ColumnName == "ObjSymb")
                            {
                                sValue = sValue.Replace(" ", "");
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
                            if (col.ColumnName == "EvatRsltLftmYear") sValue = (sValue != "" ? sValue + "년" : sValue);
                            if (col.ColumnName == "EvatRsltPasDprcRate") sValue = sValue + "%";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db4ObjRePurcAmtTot = 0;
                    double db4NumOfMachines = 0;
                    double db4InsureValue = 0;
                    double db4Amt = 0;
                    double db4ObjCostTot = 0;

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
                                if (col.ColumnName == "ObjBuyDt")
                                {
                                    if (sValue != "") sValue = Utils.Mid(sValue, 1, 4) + "년 " + Utils.Mid(sValue, 5, 6) + "월";
                                }
                                if (col.ColumnName == "ObjRePurcAmt")
                                {
                                    db4ObjRePurcAmtTot += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjArea")
                                {
                                    db4NumOfMachines += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjDprcTotRate") sValue = Utils.Round(sValue, 2) + "%";
                                if (col.ColumnName == "ObjInsureValue")
                                {
                                    db4InsureValue += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "LosAmt")
                                {
                                    db4Amt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjCost")
                                {
                                    db4ObjCostTot += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTableA.GetRow((i + 1) * 2 + 0), sKey, sValue);
                                rUtil.ReplaceTableRow(oTableA.GetRow((i + 1) * 2 + 1), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTableA.GetRow((dtB.Rows.Count + 1) * 2), "@db4NumOfMachines@", Utils.AddComma(db4NumOfMachines));
                    rUtil.ReplaceTableRow(oTableA.GetRow((dtB.Rows.Count + 1) * 2), "@db4ObjRePurcAmtTot@", Utils.AddComma(db4ObjRePurcAmtTot));
                    rUtil.ReplaceTableRow(oTableA.GetRow((dtB.Rows.Count + 1) * 2), "@db4InsureValue@", Utils.AddComma(db4InsureValue));
                    rUtil.ReplaceTableRow(oTableA.GetRow((dtB.Rows.Count + 1) * 2), "@db4Amt@", Utils.AddComma(db4Amt));
                    rUtil.ReplaceTableRow(oTableA.GetRow((dtB.Rows.Count + 1) * 2), "@db4ObjCostTot@", Utils.AddComma(db4ObjCostTot));

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
