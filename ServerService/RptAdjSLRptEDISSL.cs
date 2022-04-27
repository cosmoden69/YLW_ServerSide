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
    public class RptAdjSLRptEDISSL
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptEDISSL(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisEDISSLReport";
                security.methodId = "Query";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock1");

                dt.Columns.Add("req_id");

                dt.Clear();
                DataRow dr = dt.Rows.Add();

                dr["req_id"] = para.AcptMgmtSeq;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "Start");

                string sSampleXSD = myPath + @"\보고서\RptAdjSLRptEDISSL.xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\RptAdjSLRptEDISSL.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["contrNm"]);
                }
                rptName = "전문보고서_삼성계약적부(" + sfilename + ").docx";
                rptPath = sSample1Relt;
                //System.Diagnostics.Process process = System.Diagnostics.Process.Start(sSample1Relt);
                //Utils.BringWindowToTop(process.Handle);

                return new Response() { Result = 1, Message = "OK" };
            }
            catch (Exception ex)
            {
                return new Response() { Result = -99, Message = ex.Message };
            }
        }

        private string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
        {
            string sRet = "";

            if (!File.Exists(sDocFile)) return RptUtils.GetMessage("원본파일(word)이 존재하지 않습니다.", sDocFile);
            if (!File.Exists(sXSDFile)) return RptUtils.GetMessage("XSD파일이 존재하지 않습니다.", sXSDFile);

            DataTable dtB = null;
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
                    Table oTblImage = rUtil.GetTable(lstTable, "@B3AcdtPictImage@");

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

                    IEnumerable<Table> lstTable = doc.Body.Elements<Table>();
                    Table oTblImage = rUtil.GetTable(lstTable, "@B3AcdtPictImage@");

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "BdSurvNm")
                            {
                                sValue = sValue + (dr["BdSurvTelno"] + "" == "" ? "" : "(" + dr["BdSurvTelno"] + ")");
                            }
                            if (col.ColumnName == "BdIntrvwNm")
                            {
                                sValue = sValue + (dr["BdIntrvwRel"] + "" == "" ? "" : "(" + dr["BdIntrvwRel"] + ")");
                            }
                            if (col.ColumnName == "BdLvlChk")
                            {
                                if (sValue == "1") rUtil.ReplaceTables(lstTable, "@B2BdLvlChk@", "■");
                                else rUtil.ReplaceTables(lstTable, "@B2BdLvlChk@", "□");
                                continue;
                            }
                            if (col.ColumnName == "BdJoinBzTypeChk")
                            {
                                if (sValue == "1") rUtil.ReplaceTables(lstTable, "@B2BdJoinBzTypeChk@", "■");
                                else rUtil.ReplaceTables(lstTable, "@B2BdJoinBzTypeChk@", "□");
                                continue;
                            }
                            if (col.ColumnName == "BdRateBzTypeChk")
                            {
                                if (sValue == "1") rUtil.ReplaceTables(lstTable, "@B2BdRateBzTypeChk@", "■");
                                else rUtil.ReplaceTables(lstTable, "@B2BdRateBzTypeChk@", "□");
                                continue;
                            }
                            if (col.ColumnName == "BdAddrChk")
                            {
                                if (sValue == "1") rUtil.ReplaceTables(lstTable, "@B2BdAddrChk@", "■");
                                else rUtil.ReplaceTables(lstTable, "@B2BdAddrChk@", "□");
                                continue;
                            }
                            if (col.ColumnName == "BdownrNmChk")
                            {
                                if (sValue == "1") rUtil.ReplaceTables(lstTable, "@B2BdownrNmChk@", "■");
                                else rUtil.ReplaceTables(lstTable, "@B2BdownrNmChk@", "□");
                                continue;
                            }
                            if (col.ColumnName == "strPanYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strPanY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2strPanN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strPanY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2strPanN@", "■");
                                }
                                continue;
                            }
                            if (col.ColumnName == "strTentYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strTentY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2strTentN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strTentY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2strTentN@", "■");
                                }
                                continue;
                            }
                            if (col.ColumnName == "strFireWallYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strFireWallY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2strFireWallN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2strFireWallY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2strFireWallN@", "■");
                                }
                                continue;
                            }
                            if (col.ColumnName == "DisStoreYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2DisStoreY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2DisStoreN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2DisStoreY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2DisStoreN@", "■");
                                }
                                continue;
                            }
                            if (col.ColumnName == "TradMarketYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2TradMarketY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2TradMarketN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2TradMarketY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2TradMarketN@", "■");
                                }
                                continue;
                            }
                            if (col.ColumnName == "LocChkYN")
                            {
                                if (sValue == "1")
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2LocChkY@", "■");
                                    rUtil.ReplaceTables(lstTable, "@B2LocChkN@", "□");
                                }
                                else
                                {
                                    rUtil.ReplaceTables(lstTable, "@B2LocChkY@", "□");
                                    rUtil.ReplaceTables(lstTable, "@B2LocChkN@", "■");
                                }
                                continue;
                            }

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null && oTblImage != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        int rnum = 0;
                        int rmdr = 0;

                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                        sValue = dr["AcdtPictImage"] + "";
                        TableRow xrow1 = oTblImage.GetRow(rnum);
                        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                        try
                        {
                            Image img = Utils.stringToImage(sValue);
                            rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2730000L, 2100000L);
                        }
                        catch { }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null && oTblImage != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        int rnum = 0;
                        int rmdr = 1;

                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                        sValue = dr["AcdtPictImage"] + "";
                        TableRow xrow1 = oTblImage.GetRow(rnum);
                        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                        try
                        {
                            Image img = Utils.stringToImage(sValue);
                            rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2730000L, 2100000L);
                        }
                        catch { }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null && oTblImage != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        int rnum = 1;
                        int rmdr = 0;

                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                        sValue = dr["AcdtPictImage"] + "";
                        TableRow xrow1 = oTblImage.GetRow(rnum);
                        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                        try
                        {
                            Image img = Utils.stringToImage(sValue);
                            rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2730000L, 2100000L);
                        }
                        catch { }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null && oTblImage != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        int rnum = 1;
                        int rmdr = 1;

                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                        sValue = dr["AcdtPictImage"] + "";
                        TableRow xrow1 = oTblImage.GetRow(rnum);
                        rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                        try
                        {
                            Image img = Utils.stringToImage(sValue);
                            rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2730000L, 2100000L);
                        }
                        catch { }
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
