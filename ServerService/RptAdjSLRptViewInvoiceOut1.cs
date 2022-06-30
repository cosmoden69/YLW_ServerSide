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
{//
    public class RptAdjSLRptViewInvoiceOut1
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptViewInvoiceOut1(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisAdjSLInvoiceViewBillIssuePrint";
                security.methodId = "PrintOut";
                security.companySeq = para.CompanySeq;

                JObject jparam = JObject.Parse(para.ParamStr);

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock11");
                dt.Columns.Add("TaxUnit");
                dt.Clear();
                DataRow dr = dt.Rows.Add();
                dr["TaxUnit"] = jparam["TaxUnit"];

                dt = ds.Tables.Add("DataBlock12");
                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");
                dt.Columns.Add("AcptMgmtNo");
                dt.Columns.Add("InvcSeq");
                dt.Columns.Add("InvcSeqKeyQry");
                dt.Columns.Add("DeptHeadCd");
                dt.Columns.Add("CustSeq");
                dt.Clear();
                dr = dt.Rows.Add();
                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;
                dr["InvcSeq"] = jparam["InvcSeq"];
                dr["DeptHeadCd"] = jparam["DeptHeadCd"];

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "Start");

                string sSampleXSD = myPath + @"\보고서\출력설계_1715_서식_인보이스(외부용).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_1715_서식_인보이스(외부용).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                DataTable dtB = pds.Tables["DataBlock13"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    string fileSeq = Utils.ConvertToString(dtB.Rows[0]["CostRcptFileSeq"]);
                    DataSet pds1 = YLWService.MTRServiceModule.CallMTRFileDownload(security, fileSeq, "", "");
                    if (pds1 != null && pds1.Tables.Count > 0)
                    {
                        DataTable dtB1 = pds1.Tables[0];
                        if (dtB1 != null && dtB1.Rows.Count > 0)
                        {
                            sSampleDocx = myPath + @"\보고서\출력설계_1715_서식_인보이스(외부용)_Image.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            sRet = SetSample_Image(sSampleDocx, sSampleXSD, pds1, sSampleAddFile);
                            if (sRet != "")
                            {
                                return new Response() { Result = -1, Message = sRet };
                            }
                            addFiles.Add(sSampleAddFile);
                        }
                    }
                }

                //DOCX 파일합치기 
                WordprocessingDocument wdoc = WordprocessingDocument.Open(sSample1Relt, true);
                MainDocumentPart mainPart = wdoc.MainDocumentPart;
                for (int ii = 0; ii < addFiles.Count; ii++)
                {
                    string addFile = addFiles[ii];
                    RptUtils.AppendFile(mainPart, addFile, (ii > 0 ? true : false));
                    Utils.DeleteFile(addFile);
                }
                mainPart.Document.Save();
                wdoc.Close();

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                dtB = pds.Tables["DataBlock13"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "인보이스출력_인보험(" + sfilename + ").docx";
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
                    Table oTbl상세항목 = rUtil.GetTable(lstTable, "@B13DmndSubTotCof@");

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
                    Table oTbl상세항목 = rUtil.GetTable(lstTable, "@B13DmndSubTotCof@");
                    
                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];

                        sKey = rUtil.GetFieldName(sPrefix, "AcptDt");
                        sValue = Utils.DateFormat(dr["AcptDt"], "yyyy년 MM월 dd일");
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        sKey = rUtil.GetFieldName(sPrefix, "LasRptSbmsDt");
                        sValue = Utils.DateFormat(dr["LasRptSbmsDt"], "yyyy년 MM월 dd일");
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "InvcAdjFee") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DayExps") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "TrspExps") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "OthAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcIctvAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DocuAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InvcCsltReqAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DmndSubTot") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DmndSubTotCof") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SealPhoto")
                            {
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.ReplaceInternalImage(sKey, img);
                                }
                                catch { }
                                continue;
                            }
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
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

        private string SetSample_Image(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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

                    dtB = pds.Tables[0];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CostRcptFileImage");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableAddRow(oTable, 0, 1);
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

                    dtB = pds.Tables[0];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "CostRcptFileImage");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = i;
                                int rmdr = 0;

                                sKey = rUtil.GetFieldName(sPrefix, "CostRcptFileImage");
                                sValue = dr["FileBase64"] + "";
                                TableRow xrow1 = oTable.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 600000L, 50000L, 6000000L, 3500000L);
                                }
                                catch { }
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
