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
    public class RptAdjSLSurvMidRptGoods2NH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvMidRptGoods2NH_S(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtGoodsPrint";
                security.methodId = "QueryNH";
                security.companySeq = para.CompanySeq;

                DataSet ds = new DataSet("ROOT");
                DataTable dt = ds.Tables.Add("DataBlock1");

                dt.Columns.Add("AcptMgmtSeq");
                dt.Columns.Add("ReSurvAsgnNo");

                dt.Clear();
                DataRow dr = dt.Rows.Add();

                dr["AcptMgmtSeq"] = para.AcptMgmtSeq;   //496, 877
                dr["ReSurvAsgnNo"] = para.ReSurvAsgnNo;

                DataSet yds = YLWService.YLWServiceModule.CallYlwServiceCallPost(security, ds);
                if (yds == null)
                {
                    return new Response() { Result = -1, Message = "데이타가 없습니다" };
                }

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "Start");

                string sSampleXSD = myPath + @"\보고서\출력설계_2567_서식_농협_진행보고서(재물-대물, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2567_서식_농협_진행보고서(재물-대물, 간편).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSample1Docx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "진행보고서_재물-대물, 간편(" + sfilename + ").docx";
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
                    Table oTblA = rUtil.GetTable(lstTable, "@B3InsurObjDvs@"); // 2. 손해배상책임

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        //sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        //Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTblA != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTblA, 1, dtB.Rows.Count - 1);
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
                    
                    Table oTblA = rUtil.GetTable(lstTable, "@B3InsurObjDvs@"); // 2. 손해배상책임
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@db3ObjLosAmt@");
                   
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
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CclsExptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "VitmNglgRate")
                            {
                                if (Utils.ConvertToInt(dr["VitmNglgRate"]) != 0) { sValue = sValue + "%"; }
                            }
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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "IsrdTel") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db3ObjLosAmt = 0;
                    double db3ObjRmnAmt = 0;
                    double db3ObjSelfBearAmt = 0;
                    double db3ObjGivInsurAmt = 0;
                    string tmp = "";
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        //sKey = rUtil.GetFieldName(sPrefix, "ObjRmnAmt");
                        //Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTblA != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];

                                if (!dtB.Columns.Contains("Total_A")) dtB.Columns.Add("Total_A");
                                {
                                    dr["Total_A"] = Utils.ToDouble(dr["ObjRstrGexpTot"]) + Utils.ToDouble(dr["RePurcGexpAmt"]);
                                }

                                if (!dtB.Columns.Contains("Total_B")) dtB.Columns.Add("Total_B");
                                {
                                    dr["Total_B"] = Utils.ToDouble(dr["ObjRmnRmvTot"]) + Utils.ToDouble(dr["RmnObjRmvGexpAmt"]);
                                }

                                if (dtB.Rows.Count == 1) { oTblARow.Remove(); }
                            
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjLosAmt")
                                    {
                                        db3ObjLosAmt += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    if (col.ColumnName == "ObjRmnAmt")
                                    {
                                        db3ObjRmnAmt += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    if (col.ColumnName == "ObjSelfBearAmt")
                                    {
                                        db3ObjSelfBearAmt += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    if (col.ColumnName == "ObjGivInsurAmt")
                                    {
                                        db3ObjGivInsurAmt += Utils.ToDouble(sValue);
                                        sValue = Utils.AddComma(sValue);
                                    }
                                    
                                    if (col.ColumnName == "ObjRstrGexpTot") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RePurcGexpAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "ObjRmnRmvTot") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjRmvGexpAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "PureLosAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "Total_A") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "Total_B") sValue = Utils.AddComma(sValue);

                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTblA.GetRow(i + 1), sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblB.GetRow(i + 1), sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblC_1Row, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblC_2Row, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblC_3Row, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblF_1Row, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblF_2Row, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTblF_3Row, sKey, sValue);
                                }
                                tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"] + "/" + dr["ObjStrt"];
                            }
                        }
                    }
                    rUtil.ReplaceTables(lstTable, "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjRmnAmt@", Utils.AddComma(db3ObjRmnAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjStrtRmk@", tmp);
                    

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
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
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
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
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
