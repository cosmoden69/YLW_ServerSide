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
    public class RptAdjSLSurvMidRptLiabilityPersNH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvMidRptLiabilityPersNH_S(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintPersNH";
                security.methodId = "QueryCclsPrtLiabilityPerNH";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2575_서식_농협_진행보고서(배책-대인, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2575_서식_농협_진행보고서(배책-대인, 간편).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSample1Docx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "진행보고서_배책_대인, 간편(" + sfilename + ").docx";
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

                    

                    dtB = pds.Tables["DataBlock5"];
                    if (dtB != null)
                    {
                        sKey = "@B5FileNo@";
                        Table oTableA = rUtil.GetTable(lstTable, sKey);
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableA, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        sKey = "@B10PrgMgtDt@";
                        Table oTableB = rUtil.GetTable(lstTable, sKey);
                        if (oTableB != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableB, 1, dtB.Rows.Count - 1);
                        }
                    }

                    //피해자
                    dtB = pds.Tables["DataBlock12"];
                    if (dtB != null)
                    {
                        sKey = "@B12VitmNm@";
                        Table oTableC = rUtil.GetTable(lstTable, sKey);
                        if (oTableC != null)
                        {
                            //테이블의 끝에 추가
                            dtB.Rows.Count.Equals(dtB.Columns);
                            rUtil.TableInsertRows(oTableC, 2, 2, dtB.Rows.Count - 1);
                            //rUtil.TableInsertRows(oTableC, (dtB.Rows.Count * 2) + 4, 3, dtB.Rows.Count - 1);
                        }
                    }

                    //기왕력검토
                    int B12RowsCount = dtB.Rows.Count;
                    dtB = pds.Tables["DataBlock4"];
                    if (dtB != null)
                    {
                        sKey = "@B4MedHstr@";
                        Table oTableD = rUtil.GetTable(lstTable, sKey);
                        if (oTableD != null)
                        {
                            //테이블의 끝에 추가
                            //rUtil.TableInsertRows(oTableC, 2, 2, dtB.Rows.Count - 1);
                            //rUtil.TableInsertRows(oTableC, (dtB.Rows.Count * 2) + 4, 3, dtB.Rows.Count - 1);
                            rUtil.TableInsertRows(oTableD, (B12RowsCount * 2) + 4, 3, dtB.Rows.Count - 1);
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
                    sKey = "@B1DiMedfeeTot@";
                    Table oTblA = rUtil.GetTable(lstTable, sKey); //추정지급보험금 표
                    Table oTbl손해배상책임 = rUtil.GetTable(lstTable, "@B3LegaRspsbFg@");

                    sKey = "@B1DiOthExpsAmt1@"; //기타손해배상금1 행
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    sKey = "@B1DiOthExpsAmt2@"; //기타손해배상금2 행
                    TableRow oTblBRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    sKey = "@B1DiOthExpsAmt3@"; //기타손해배상금3 행
                    TableRow oTblCRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    sKey = "@B1DiOthExpsAmt4@"; //기타손해배상금4 행
                    TableRow oTblDRow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    sKey = "@B1DiOthExpsAmt5@"; //기타손해배상금5 행
                    TableRow oTblERow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                   

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        //치료비
                        if (!dtB.Columns.Contains("DiMedfeeTot")) dtB.Columns.Add("DiMedfeeTot");
                        {
                            dr["DiMedfeeTot"] = Utils.AddComma(Utils.ToDouble(dr["DiMedfeeOpt"]) + Utils.ToDouble(dr["DiMedfeeInHosp"]));
                        }

                        //향후치료비
                        if (!dtB.Columns.Contains("DiNxtMedfeeTot")) dtB.Columns.Add("DiNxtMedfeeTot");
                        {
                            dr["DiNxtMedfeeTot"] = Utils.AddComma(Utils.ToDouble(dr["DiNxtMedfee1"]) + Utils.ToDouble(dr["DiNxtMedfee2"]) +
                                                                  Utils.ToDouble(dr["DiNxtMedfee3"]) + Utils.ToDouble(dr["DiNxtMedfee4"]));
                        }

                        //소계-손해액
                        if (!dtB.Columns.Contains("SubTotA")) dtB.Columns.Add("SubTotA");
                        {
                            dr["SubTotA"] = Utils.AddComma(Utils.ToDouble(dr["DiMedfeeTot"]) + Utils.ToDouble(dr["DiNxtMedfeeTot"]) + Utils.ToDouble(dr["DiShdnLosAmt"]) +
                                                           Utils.ToDouble(dr["DiLosPrfAmt"]) + Utils.ToDouble(dr["DiNursAmt"]) + Utils.ToDouble(dr["DiOthExpsAmt1"]) +
                                                           Utils.ToDouble(dr["DiOthExpsAmt2"]) + Utils.ToDouble(dr["DiOthExpsAmt3"]) + Utils.ToDouble(dr["DiOthExpsAmt4"]) +
                                                           Utils.ToDouble(dr["DiOthExpsAmt5"]));
                        }

                        //과실상계 후 금액
                        if (!dtB.Columns.Contains("SubTotB")) dtB.Columns.Add("SubTotB");
                        {
                            dr["SubTotB"] = Utils.AddComma(Utils.ToDouble(dr["SubTotA"]) - Utils.ToDouble(dr["DiNglgBearAmt"]));
                        }

                        //기타손해배상금 명
                        if (!dtB.Columns.Contains("DiOthExpsHed1T")) dtB.Columns.Add("DiOthExpsHed1T");
                        {
                            if (Utils.ConvertToString(dr["DiOthExpsHed1"]) == "")
                            {
                                dr["DiOthExpsHed1T"] = "기타 손해1";
                            }
                            else
                            {
                                dr["DiOthExpsHed1T"] = dr["DiOthExpsHed1"];
                            }
                        }

                        if (!dtB.Columns.Contains("DiOthExpsHed2T")) dtB.Columns.Add("DiOthExpsHed2T");
                        {
                            if (Utils.ConvertToString(dr["DiOthExpsHed2"]) == "")
                            {
                                dr["DiOthExpsHed2T"] = "기타 손해2";
                            }
                            else
                            {
                                dr["DiOthExpsHed2T"] = dr["DiOthExpsHed2"];
                            }
                        }

                        if (!dtB.Columns.Contains("DiOthExpsHed3T")) dtB.Columns.Add("DiOthExpsHed3T");
                        {
                            if (Utils.ConvertToString(dr["DiOthExpsHed3"]) == "")
                            {
                                dr["DiOthExpsHed3"] = "기타 손해3";
                            }
                            else
                            {
                                dr["DiOthExpsHed3T"] = dr["DiOthExpsHed3"];
                            }
                        }

                        if (!dtB.Columns.Contains("DiOthExpsHed4T")) dtB.Columns.Add("DiOthExpsHed4T");
                        {
                            if (Utils.ConvertToString(dr["DiOthExpsHed4"]) == "")
                            {
                                dr["DiOthExpsHed4T"] = "기타 손해2";
                            }
                            else
                            {
                                dr["DiOthExpsHed4T"] = dr["DiOthExpsHed4"];
                            }
                        }

                        if (!dtB.Columns.Contains("DiOthExpsHed5T")) dtB.Columns.Add("DiOthExpsHed5T");
                        {
                            if (Utils.ConvertToString(dr["DiOthExpsHed5"]) == "")
                            {
                                dr["DiOthExpsHed5T"] = "기타 손해5";
                            }
                            else
                            {
                                dr["DiOthExpsHed2T"] = dr["DiOthExpsHed2"];
                            }
                        }


                        //기타손해배상금1~5 값이 없으면 행 지우기
                        string DiOthExpsAmt1 = Utils.AddComma(dr["DiOthExpsAmt1"]);
                        string DiOthExpsAmt2 = Utils.AddComma(dr["DiOthExpsAmt2"]);
                        string DiOthExpsAmt3 = Utils.AddComma(dr["DiOthExpsAmt3"]);
                        string DiOthExpsAmt4 = Utils.AddComma(dr["DiOthExpsAmt4"]);
                        string DiOthExpsAmt5 = Utils.AddComma(dr["DiOthExpsAmt5"]);
                        //if ((DiNxtMedReq3str == "" || DiNxtMedReq3str == "0") && (DiNxtMedfee3str == "" || DiNxtMedfee3str == "0")) { oTblBRow.Remove(); }
                        if (DiOthExpsAmt1 == "0") { oTblARow.Remove(); }
                        if (DiOthExpsAmt2 == "0") { oTblBRow.Remove(); }
                        if (DiOthExpsAmt3 == "0") { oTblCRow.Remove(); }
                        if (DiOthExpsAmt4 == "0") { oTblDRow.Remove(); }
                        if (DiOthExpsAmt5 == "0") { oTblERow.Remove(); }

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
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            //if (col.ColumnName == "ObjSelfBearAmt") sValue = Utils.AddComma(sValue);


                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            //4. 추정지급보험금 표
                            if (col.ColumnName == "DiMedReqOpt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiMedfeeOpt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiMedReqInHosp") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiMedfeeInHosp") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiShdnLosReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiShdnLosAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiLosPrfAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiNursAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiOthExpsAmt1") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiOthExpsAmt2") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiOthExpsAmt3") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiOthExpsAmt4") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiOthExpsAmt5") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiNglgBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSltmAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DIGivInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SealPhoto" || col.ColumnName == "ChrgAdjPhoto" || col.ColumnName == "LeadAdjPhoto")
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
                            if (col.ColumnName == "VitmNglgRate") sValue = sValue + "%";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    

                    //dtB = pds.Tables["DataBlock4"];
                    //sPrefix = "B4";
                    //if (dtB != null)
                    //{
                    //    if (dtB.Rows.Count < 1) dtB.Rows.Add();
                    //    DataRow dr = dtB.Rows[0];

                    //    foreach (DataColumn col in dtB.Columns)
                    //    {
                    //        sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //        sValue = dr[col] + "";
                    //        if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //        if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                    //        rUtil.ReplaceTables(lstTable, sKey, sValue);
                    //    }
                    //}

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
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
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 2), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        double db6ObjSelfBearAmt = 0; //자기부담금
                        string tmp = "";

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjSelfBearAmt") db6ObjSelfBearAmt += Utils.ToDouble(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                        rUtil.ReplaceTables(lstTable, "@db6ObjSelfBearAmt@", Utils.AddComma(db6ObjSelfBearAmt));

                        tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"] + "/" + dr["ObjStrt"];
                        rUtil.ReplaceTables(lstTable, "@db6ObjStrtRmk@", tmp);
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
                                //DataRow dr = dtB.Rows[0];

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

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "VitmNm");
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
                                    if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "VitmSubSeq") {
                                        if (dtB.Rows.Count == 1)
                                        {
                                            sValue = "";
                                        }
                                    }
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    //rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow((i + 1) * 2 + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow((i + 1) * 2 + 1), sKey, sValue);

                                }
                            }
                        }
                    }
                    int B12Count = dtB.Rows.Count;
                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "MedHstr");
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
                                    if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "VitmSubSeq")
                                    {
                                        if (B12Count == 1)
                                        {
                                            sValue = "";
                                        }
                                        else
                                        {
                                            sValue += "-";
                                        }
                                    }
                                    if (col.ColumnName == "CureSeq")
                                    {
                                        if (dtB.Rows.Count == 1)
                                        {
                                            sValue = "";
                                        }
                                    }
                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 1), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTable.GetRow(((i * 3) + 1) + (B12Count * 2) + 3 + 2), sKey, sValue);
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
