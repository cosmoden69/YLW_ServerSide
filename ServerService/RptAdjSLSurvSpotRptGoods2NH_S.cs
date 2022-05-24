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
    public class RptAdjSLSurvSpotRptGoods2NH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvSpotRptGoods2NH_S(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2568_서식_농협_현장보고서(재물-대물, 간편).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_2568_서식_농협_현장보고서(재물-대물, 간편).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                DataTable dtBT = pds.Tables["DataBlock14"];
                if (dtBT != null && dtBT.Rows.Count > 0)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2568_서식_농협_현장보고서(재물-대물, 간편)_Pict.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    sRet = SetSample1Pict(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }

                //DOCX 파일합치기 
                WordprocessingDocument wdoc = WordprocessingDocument.Open(sSample1Relt, true);
                MainDocumentPart mainPart = wdoc.MainDocumentPart;
                for (int ii = 0; ii < addFiles.Count; ii++)
                {
                    string addFile = addFiles[ii];
                    RptUtils.AppendFile(mainPart, addFile, true);
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
                DataTable dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "현장보고서_재물-대물, 간편(" + sfilename + ").docx";
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
                    //Table oTblA = rUtil.GetTable(lstTable, "@B3ObjRmnAmt@"); // 1. 총괄표
                    Table oTblA = rUtil.GetTable(lstTable, "@B3InsurObjDvs@"); // 2. 손해배상책임
                    //Table oTblC = rUtil.GetTable(lstTable, "@B3RePurcGexpAmt@"); // 5. 세부평가내역
                    //Table oTblD = rUtil.GetTable(lstTable, "@B8RmnObjCost@"); // 6. 잔존물 및 구상 표1
                    //Table oTblE = rUtil.GetTable(lstTable, "@B8SucBidDt@"); // 6. 잔존물 및 구상 표2
                    //Table oTblF = rUtil.GetTable(lstTable, "@B3RmnObjRmvGexpAmt@"); // 6. 잔존물 및 구상 - 잔존물 및 제거비용

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

                    //dtB = pds.Tables["DataBlock3"];
                    //sPrefix = "B3";
                    //if (dtB != null)
                    //{
                    //    if (oTblA != null)
                    //    {
                    //        //테이블의 중간에 추가
                    //        rUtil.TableInsertRow(oTblA, 1, dtB.Rows.Count - 1);
                    //    }

                    //    if (oTblB != null)
                    //    {
                    //        //테이블의 중간에 추가
                    //        rUtil.TableAddRow(oTblB, 1, dtB.Rows.Count - 1);
                    //    }
                    //}

                    ////5.세부평가내역
                    //drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 2");
                    //if (drs == null || drs.Length < 1)
                    //{
                    //    if (oTblC != null) rUtil.TableRemoveRow(oTblC, 1);
                    //}
                    //else
                    //{
                    //    if (oTblC != null)
                    //    {
                    //        //테이블의 중간에 삽입
                    //        rUtil.TableInsertRow(oTblC, 1, drs.Length - 1);
                    //    }
                    //}

                    ////6. 잔존물 및 구상 표1
                    //drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    //if (drs == null || drs.Length < 1)
                    //{
                    //    //if (oTblD != null) rUtil.TableRemoveRow(oTblD, 1);
                    //    //if (oTblD != null) oTblD.Remove();
                    //}
                    //else
                    //{
                    //    if (oTblD != null)
                    //    {
                    //        //테이블의 중간에 삽입
                    //        rUtil.TableInsertRow(oTblD, 1, drs.Length - 1);
                    //    }
                    //}

                    ////6. 잔존물 및 구상 표2
                    //drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    //if (drs == null || drs.Length < 1)
                    //{
                    //    //if (oTblD != null) rUtil.TableRemoveRow(oTblD, 1);
                    //    //if (oTblE != null) oTblE.Remove();
                    //}
                    //else
                    //{
                    //    if (oTblE != null)
                    //    {
                    //        //테이블의 중간에 삽입
                    //        rUtil.TableInsertRow(oTblE, 1, drs.Length - 1);
                    //    }
                    //}

                    ////6. 잔존물 및 구상 - 잔존물 및 제거비용
                    //drs = pds.Tables["DataBlock5"]?.Select("EvatCd % 10 = 3");
                    //if (drs == null || drs.Length < 1)
                    //{
                    //    if (oTblF != null) rUtil.TableRemoveRow(oTblF, 1);
                    //}
                    //else
                    //{
                    //    if (oTblF != null)
                    //    {
                    //        //테이블의 중간에 삽입
                    //        rUtil.TableInsertRow(oTblF, 1, drs.Length - 1);
                    //    }
                    //}

                    //dtB = pds.Tables["DataBlock9"];
                    //sPrefix = "B9";
                    //if (dtB != null)
                    //{
                    //    sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                    //    Table oTable = rUtil.GetTable(lstTable, sKey);
                    //    if (oTable != null)
                    //    {
                    //        //테이블의 끝에 추가
                    //        rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                    //    }
                    //}





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

                    //Table oTblA = rUtil.GetTable(lstTable, "@B3ObjRmnAmt@"); // 1. 총괄표
                    //TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@db3ObjLosAmt@");
                    Table oTbl표지 = rUtil.GetTable(lstTable, "@B1ChrgAdjManRegNo@");
                    Table oTblA = rUtil.GetTable(lstTable, "@B3InsurObjDvs@"); // 2. 손해배상책임
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@db3ObjLosAmt@");

                    //Table oTblC = rUtil.GetTable(lstTable, "@B3RePurcGexpAmt@"); // 5. 세부평가내역
                    //TableRow oTblC_1Row = rUtil.GetTableRow(oTblC?.Elements<TableRow>(), "@B3ObjRstrGexpTot@");
                    //TableRow oTblC_2Row = rUtil.GetTableRow(oTblC?.Elements<TableRow>(), "@B3RstrGexpRate@");
                    //TableRow oTblC_3Row = rUtil.GetTableRow(oTblC?.Elements<TableRow>(), "@B3Total_A@");
                    //Table oTblD = rUtil.GetTable(lstTable, "@B8RmnObjNm@"); // 6. 잔존물 및 구상 표1
                    //Table oTblE = rUtil.GetTable(lstTable, "@B8SucBidDt@"); // 6. 잔존물 및 구상 표2
                    //Table oTblF = rUtil.GetTable(lstTable, "@B3RmnObjRmvGexpAmt@"); // 6. 잔존물 및 구상 - 잔존물제거비용
                    //TableRow oTblF_1Row = rUtil.GetTableRow(oTblF?.Elements<TableRow>(), "@B3ObjRmnRmvTot@");
                    //TableRow oTblF_2Row = rUtil.GetTableRow(oTblF?.Elements<TableRow>(), "@B3RmnObjRmvGexpAmt@");
                    //TableRow oTblF_3Row = rUtil.GetTableRow(oTblF?.Elements<TableRow>(), "@B3Total_B@");

                    var db1SurvAsgnEmpManRegNo = ""; //조사자 손해사정등록번호
                    var db1SurvAsgnEmpAssRegNo = ""; //조사자 보조인 등록번호
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
                            if (col.ColumnName == "VitmNglgRate")
                            {
                                if (Utils.ConvertToInt(dr["VitmNglgRate"]) != 0) { sValue = sValue + "%"; }
                            }
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
                            if (col.ColumnName == "LeadAdjManRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "ChrgAdjManRegNo")
                            {
                                if (sValue != "") sValue = "손해사정등록번호 : 제" + sValue + "호";
                            }
                            if (col.ColumnName == "SurvAsgnEmpManRegNo")
                            {
                                if (sValue != "") db1SurvAsgnEmpManRegNo = sValue;
                            }
                            if (col.ColumnName == "SurvAsgnEmpAssRegNo")
                            {
                                if (sValue != "") db1SurvAsgnEmpAssRegNo = sValue;
                            }

                            rUtil.ReplaceHeaderPart(doc, sKey, sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }
                    if (db1SurvAsgnEmpManRegNo == "")
                    {
                        if (db1SurvAsgnEmpAssRegNo != "")
                        {
                            db1SurvAsgnEmpAssRegNo = "보조인 등록번호 : 제" + db1SurvAsgnEmpAssRegNo + "호";
                        }
                        rUtil.ReplaceTable(oTbl표지, "@db1SurvAsgnEmpRegNo@", db1SurvAsgnEmpAssRegNo);
                    }
                    else
                    {
                        db1SurvAsgnEmpManRegNo = "손해사정등록번호 : 제" + db1SurvAsgnEmpManRegNo + "호";
                        rUtil.ReplaceTable(oTbl표지, "@db1SurvAsgnEmpRegNo@", db1SurvAsgnEmpManRegNo);
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
                                if (Utils.ConvertToString(dr["InsurObjDvs"]) != "")
                                {
                                    tmp += "\n" + dr["InsurObjDvs"] + "/" + dr["ObjStrt"];
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTables(lstTable, "@db3ObjLosAmt@", Utils.AddComma(db3ObjLosAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjRmnAmt@", Utils.AddComma(db3ObjRmnAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjSelfBearAmt@", Utils.AddComma(db3ObjSelfBearAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjGivInsurAmt@", Utils.AddComma(db3ObjGivInsurAmt));
                    rUtil.ReplaceTables(lstTable, "@db3ObjStrtRmk@", tmp);
                    

                    //dtB = pds.Tables["DataBlock5"];
                    //sPrefix = "B5";
                    //if (dtB != null)
                    //{
                    //    int ia = 0, ib = 0;
                    //    for (int i = 0; i < dtB.Rows.Count; i++)
                    //    {
                    //        DataRow dr = dtB.Rows[i];
                    //        int EvatCd = Utils.ToInt(dtB.Rows[i]["EvatCd"]);

                    //        if (EvatCd % 10 == 2)  // 5. 세부평가내역
                    //        {
                    //            foreach (DataColumn col in dtB.Columns)
                    //            {
                    //                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //                sValue = dr[col] + "";
                    //                if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                    //                rUtil.ReplaceTableRow(oTblC.GetRow(ia + 1), sKey, sValue);
                    //            }
                    //            ia++;
                    //        }
                    //        if (EvatCd % 10 == 3)  // 6. 잔존물 및 구상 - 잔존물제거비용
                    //        {
                    //            foreach (DataColumn col in dtB.Columns)
                    //            {
                    //                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //                sValue = dr[col] + "";
                    //                if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                    //                rUtil.ReplaceTableRow(oTblF.GetRow(ib + 1), sKey, sValue);
                    //            }
                    //            ib++;
                    //        }
                    //        //else if (EvatCd % 10 == 3)  //잔존물제거비용
                    //        //{
                    //        //    foreach (DataColumn col in dtB.Columns)
                    //        //    {
                    //        //        sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //        //        sValue = dr[col] + "";
                    //        //        if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                    //        //        rUtil.ReplaceTableRow(oTableC.GetRow(ic + 1), sKey, sValue);
                    //        //    }
                    //        //    ic++;
                    //        //}
                    //    }
                    //}

                    //dtB = pds.Tables["DataBlock8"];
                    //sPrefix = "B8";
                    //if (dtB != null)
                    //{
                    //    if (dtB.Rows.Count < 1) dtB.Rows.Add();
                    //    {
                    //        int ic = 0, id = 0;
                    //        for (int i = 0; i < dtB.Rows.Count; i++)
                    //        {
                    //            DataRow dr = dtB.Rows[i];
                    //            int TrtCd = Utils.ToInt(dtB.Rows[i]["TrtCd"]);

                    //            if (TrtCd % 10 == 1 || TrtCd == 0)  //잔존물제거비용 표1
                    //            {
                    //                foreach (DataColumn col in dtB.Columns)
                    //                {
                    //                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //                    sValue = dr[col] + "";
                    //                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                    //                    rUtil.ReplaceTableRow(oTblD.GetRow(ic + 1), sKey, sValue);
                    //                }
                    //                ic++;
                    //            }
                    //            if (TrtCd % 10 == 2 || TrtCd == 0)  //잔존물제거비용 표2
                    //            {
                    //                foreach (DataColumn col in dtB.Columns)
                    //                {
                    //                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                    //                    sValue = dr[col] + "";
                    //                    if (col.ColumnName == "AuctFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //                    if (col.ColumnName == "AuctToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //                    if (col.ColumnName == "SucBidDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                    //                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                    //                    rUtil.ReplaceTableRow(oTblE.GetRow(id + 1), sKey, sValue);
                    //                }
                    //                id++;
                    //            }
                    //        }
                    //    }
                    //}

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

        private string SetSample1Pict(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B14AcdtPictImage@");

                    dtB = pds.Tables["DataBlock14"];
                    if (dtB != null)
                    {
                        if (oTbl현장사진 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl현장사진, 0, 1);
                                rUtil.TableAddRow(oTbl현장사진, 1, 1);
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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B14AcdtPictImage@");

                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (oTbl현장사진 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2;
                                int rmdr = i % 2;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl현장사진.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 3000000L, 2400000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl현장사진.GetRow(rnum + 1);
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
