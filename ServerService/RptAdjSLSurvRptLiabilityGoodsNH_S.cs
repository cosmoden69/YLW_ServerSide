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
    public class RptAdjSLSurvRptLiabilityGoodsNH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityGoodsNH_S(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintGoodsNH";
                security.methodId = "Query";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2574_서식_농협_종결보고서(배책-차량, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2574_서식_농협_종결보고서(배책-차량, 간편).docx";
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
                rptName = "종결보고서_배책_대물_간편(" + sfilename + ").docx";
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

                    dtB = pds.Tables["DataBlock6"];
                    if (dtB != null)
                    {
                        Table oTbl총괄표 = rUtil.GetTable(lstTable, "피해물 구분");
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl총괄표, 1, dtB.Rows.Count - 1);
                        }
                        Table oTbl피해품목 = rUtil.GetTable(lstTable, "피해품목");
                        if (oTbl피해품목 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl피해품목, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        Table oTbl보험금지급처 = rUtil.GetTable(lstTable, "@B7InsurGivObj@");
                        if (oTbl보험금지급처 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl보험금지급처, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");
                        if (oTbl진행사항 != null)
                        {
                            rUtil.TableAddRow(oTbl진행사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    drs = pds.Tables["DataBlock13"]?.Select("EvatCd % 10 = 3");
                    if (drs != null && drs.Length > 0)
                    {
                        Table oTbl잔존물 = rUtil.GetTable(lstTable, "@B13EvatCatg@");
                        if (oTbl잔존물 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl잔존물, 1, drs.Length - 1);
                        }
                    }

                    drs = pds.Tables["DataBlock14"]?.Select("TrtCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        Table oTbl당사 = rUtil.GetTable(lstTable, "당사", "@B14RmnObjNm@");
                        if (oTbl당사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl당사, 1, drs.Length - 1);
                        }
                    }
                    drs = pds.Tables["DataBlock14"]?.Select("TrtCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        Table oTbl옥션 = rUtil.GetTable(lstTable, "옥션", "@B14RmnObjNm@");
                        if (oTbl옥션 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    if (dtB != null)
                    {
                        Table oTbl별첨 = rUtil.GetTable(lstTable, "@B5FileNo@");
                        if (oTbl별첨 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl별첨, 1, dtB.Rows.Count - 1);
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
                    Table oTbl표지 = rUtil.GetTable(lstTable, "@B1LeadAdjuster@");
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "피해물 구분");
                    Table oTbl피해품목 = rUtil.GetTable(lstTable, "피해품목");
                    Table oTbl보험금지급처 = rUtil.GetTable(lstTable, "@B7InsurGivObj@");
                    Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");
                    Table oTbl잔존물 = rUtil.GetTable(lstTable, "@B13EvatCatg@");
                    Table oTbl당사 = rUtil.GetTable(lstTable, "당사", "@B14RmnObjNm@");
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "옥션", "@B14RmnObjNm@");
                    Table oTbl별첨 = rUtil.GetTable(lstTable, "@B5FileNo@");

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
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "DoFixAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoFixCmnt" && dr["DoFixBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoFixBss"];
                            }
                            if (col.ColumnName == "DoNoCarfeeCmnt" && dr["DoNoCarfeeBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoNoCarfeeBss"];
                            }
                            if (col.ColumnName == "DoRentCarCmnt" && dr["DoRentCarBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoRentCarBss"];
                            }
                            if (col.ColumnName == "DoOthExpsCmnt" && dr["DoOthExpsBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoOthExpsBss"];
                            }
                            if (col.ColumnName == "DoNglgBearCmnt" && dr["DoNglgBearBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoNglgBearBss"];
                            }
                            if (col.ColumnName == "DoSelfBearCmnt" && dr["DoSelfBearBss"] + "" != "")
                            {
                                sValue += (sValue != "" ? "\n" : "") + dr["DoSelfBearBss"];
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

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "FixFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FixToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "RentFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "RentToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl별첨 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTbl별첨.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        string tmp = "";
                        double db6ObjTotAmt = 0;
                        double db6ObjSelfBearAmt = 0;
                        double db6ObjGivInsurAmt = 0;
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"] + "/" + dr["ObjStrtRmk"];
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "ObjLosAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "ObjTotAmt")
                                {
                                    db6ObjTotAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjSelfBearAmt")
                                {
                                    db6ObjSelfBearAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjGivInsurAmt")
                                {
                                    db6ObjGivInsurAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl피해품목.GetRow(i + 1), sKey, sValue);
                            }

                        }
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjTotAmt@", Utils.AddComma(db6ObjTotAmt));
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjSelfBearAmt@", Utils.AddComma(db6ObjSelfBearAmt));
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjGivInsurAmt@", Utils.AddComma(db6ObjGivInsurAmt));
                        rUtil.ReplaceTables(lstTable, "@db6ObjStrtRmk@", tmp);
                        rUtil.ReplaceTables(lstTable, "@db6ObjSelfBearAmt@", Utils.AddComma(db6ObjSelfBearAmt));
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTbl보험금지급처 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl보험금지급처.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
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
                                if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                rUtil.ReplaceTableRow(oTbl진행사항.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (oTbl잔존물 != null)
                        {
                            double db12RmnObjRmvTot = 0;
                            double db12RmnObjRate = 0;
                            double db12RmnObjRmvAmt = 0;
                            double db12RmnObjTot = 0;
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ObjRmnRmvTot") db12RmnObjRmvTot += Utils.ToDouble(sValue);
                                    if (col.ColumnName == "RmnObjRmvGexpRate") db12RmnObjRate += Utils.ToDouble(sValue);
                                    if (col.ColumnName == "RmnObjRmvGexpAmt") db12RmnObjRmvAmt += Utils.ToDouble(sValue);
                                }
                            }
                            db12RmnObjTot = db12RmnObjRmvTot + db12RmnObjRmvAmt;
                            TableRow oTblRow = rUtil.GetTableRow(oTbl잔존물?.Elements<TableRow>(), "@db12RmnObjRmvTot@");
                            rUtil.ReplaceTableRow(oTblRow, "@db12RmnObjRmvTot@", Utils.AddComma(db12RmnObjRmvTot));
                            oTblRow = rUtil.GetTableRow(oTbl잔존물?.Elements<TableRow>(), "@db12RmnObjRmvAmt@");
                            rUtil.ReplaceTableRow(oTblRow, "@db12RmnObjRate@", Utils.AddComma(db12RmnObjRate));
                            rUtil.ReplaceTableRow(oTblRow, "@db12RmnObjRmvAmt@", Utils.AddComma(db12RmnObjRmvAmt));
                            oTblRow = rUtil.GetTableRow(oTbl잔존물?.Elements<TableRow>(), "@db12RmnObjTot@");
                            rUtil.ReplaceTableRow(oTblRow, "@db12RmnObjTot@", Utils.AddComma(db12RmnObjTot));
                        }
                    }

                    drs = pds.Tables["DataBlock13"]?.Select("EvatCd % 10 = 3");
                    sPrefix = "B13";
                    if (drs != null)
                    {
                        if (oTbl잔존물 != null)
                        {
                            if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "EvatAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl잔존물.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    drs = pds.Tables["DataBlock14"]?.Select("TrtCd % 10 = 1");
                    sPrefix = "B14";
                    if (drs != null)
                    {
                        if (oTbl당사 != null)
                        {
                            if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock14"].Rows.Add() };
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl당사.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl당사.Remove();
                    }

                    drs = pds.Tables["DataBlock14"]?.Select("TrtCd % 10 = 2");
                    sPrefix = "B14";
                    if (drs != null)
                    {
                        if (oTbl옥션 != null)
                        {
                            if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock14"].Rows.Add() };
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "AuctFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "AuctToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "SucBidDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTbl옥션.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl옥션.Remove();
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
