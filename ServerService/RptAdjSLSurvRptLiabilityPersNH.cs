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
    public class RptAdjSLSurvRptLiabilityPersNH
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityPersNH(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2571_서식_농협_종결보고서(배책-대인).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2571_서식_농협_종결보고서(배책-대인).docx";
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
                rptName = "종결보고서_배책_대인(" + sfilename + ").docx";
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

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        sKey = "보험금 지급처";
                        Table oTblA = rUtil.GetTable(lstTable, sKey);
                        sKey = "@B7InsurGivObj@";
                        TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                        Table oTableA = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableA, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock2"];
                    if (dtB != null)
                    {
                        sKey = "@B2AcdtPictImage@";
                        Table oTableB = rUtil.GetTable(lstTable, sKey);
                        if (oTableB != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableB, 1, 1);
                                rUtil.TableAddRow(oTableB, 2, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    if (dtB != null)
                    {
                        sKey = "평가 기준";
                        Table oTblC = rUtil.GetTable(lstTable, sKey);
                        sKey = "@B4VstHosp@";
                        TableRow oTblCRow = rUtil.GetTableRow(oTblC?.Elements<TableRow>(), sKey);
                        Table oTableC = oTblCRow?.GetCell(1).Elements<Table>().FirstOrDefault();
                        if (oTableC != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTableC, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPrsCcndGrp");
                        Table oTableD = rUtil.GetTable(lstTable, sKey);
                        if (oTableD != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableD, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    if (dtB != null)
                    {
                        sKey = "@B5FileNo@";
                        Table oTableE = rUtil.GetTable(lstTable, sKey);
                        if (oTableE != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableE, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    if (dtB != null)
                    {
                        sKey = "@B9AcdtPictImage@";
                        Table oTableF = rUtil.GetTable(lstTable, sKey);
                        if (oTableF != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = dtB.Rows.Count;
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableF, 0, 1);
                                rUtil.TableAddRow(oTableF, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        sKey = "@B10PrgMgtDt@";
                        Table oTableG = rUtil.GetTable(lstTable, sKey);
                        if (oTableG != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableG, 1, dtB.Rows.Count - 1);
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
                    sKey = "보험금 지급처";
                    Table oTblA = rUtil.GetTable(lstTable, sKey);
                    sKey = "@B7InsurGivObj@";
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    Table oTableA = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    sKey = "@B2AcdtPictImage@";
                    Table oTableB = rUtil.GetTable(lstTable, sKey);

                    sKey = "평가 기준";
                    Table oTblC = rUtil.GetTable(lstTable, sKey);
                    sKey = "@B4VstHosp@";
                    TableRow oTblCRow = rUtil.GetTableRow(oTblC?.Elements<TableRow>(), sKey);
                    Table oTableC = oTblCRow?.GetCell(1).Elements<Table>().FirstOrDefault();

                    sKey = "평가 결과";
                    Table oTblCA = rUtil.GetTable(lstTable, sKey);

                    sKey = "@B11AcdtPrsCcndGrp@";
                    Table oTableD = rUtil.GetTable(lstTable, sKey);

                    sKey = "@B5FileNo@";
                    Table oTableE = rUtil.GetTable(lstTable, sKey);

                    sKey = "@B9AcdtPictImage@";
                    Table oTableF = rUtil.GetTable(lstTable, sKey);

                    sKey = "@B10PrgMgtDt@";
                    Table oTableG = rUtil.GetTable(lstTable, sKey);

                    Table oTableH = rUtil.GetTable(lstTable, "@B13ExpsLosAmt92@");

                    string strAcdtDt = "";

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        strAcdtDt = dr["AcdtDt"] + "";
                        //총괄표
                        if (!dtB.Columns.Contains("db1InsurRegsAmt")) dtB.Columns.Add("db1InsurRegsAmt");
                        {
                            dr["db1InsurRegsAmt"] = Utils.AddComma(Utils.ToDouble(dr["InsurRegsAmt"]));
                        }
                        if (!dtB.Columns.Contains("db1DiTotAmt")) dtB.Columns.Add("db1DiTotAmt");
                        {
                            dr["db1DiTotAmt"] = Utils.AddComma(Utils.ToDouble(dr["DiTotAmt"]));
                        }
                        if (!dtB.Columns.Contains("db1DiSelfBearAmt")) dtB.Columns.Add("db1DiSelfBearAmt");
                        {
                            dr["db1DiSelfBearAmt"] = Utils.AddComma(Utils.ToDouble(dr["DiSelfBearAmt"]));
                        }
                        if (!dtB.Columns.Contains("db1DIGivInsurAmt")) dtB.Columns.Add("db1DIGivInsurAmt");
                        {
                            dr["db1DIGivInsurAmt"] = Utils.AddComma(Utils.ToDouble(dr["DIGivInsurAmt"]));
                        }
                        //총괄표

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
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjusterr") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DiSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DIGivInsurAmt") sValue = Utils.AddComma(sValue);
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
                        rUtil.ReplaceTables(lstTable, "@db1InsurRegsAmt@", dr["db1InsurRegsAmt"] + "");
                        rUtil.ReplaceTables(lstTable, "@db1DiTotAmt@", dr["db1DiTotAmt"] + "");
                        rUtil.ReplaceTables(lstTable, "@db1DiSelfBearAmt@", dr["db1DiSelfBearAmt"] + "");
                        rUtil.ReplaceTables(lstTable, "@db1DIGivInsurAmt@", dr["db1DIGivInsurAmt"] + "");
                    }

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (oTableB != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                                int rmdr = i % 2;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableB.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableB.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        string db3RgtCpstOpni = "";
                        if (dr["RgtCpstOpni"] + "" != "") db3RgtCpstOpni += (db3RgtCpstOpni != "" ? "\n" : "") + dr["RgtCpstOpni"];
                        if (dr["RgtCpstOthOpni"] + "" != "") db3RgtCpstOpni += (db3RgtCpstOpni != "" ? "\n" : "") + dr["RgtCpstOthOpni"];
                        if (dr["RgtCpstSrc"] + "" != "") db3RgtCpstOpni += (db3RgtCpstOpni != "" ? "\n" : "") + dr["RgtCpstSrc"];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "VitmNglgRate") sValue = sValue + "%";
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                        rUtil.ReplaceTables(lstTable, "@db3RgtCpstOpni@", db3RgtCpstOpni);
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        string db4DgnsNm = "";
                        string db4VstHosp = "";
                        string db4CureFrDt = "";
                        string db4CureCnts = "";
                        string db4MedHstr = "";
                        double db4Medfee = 0;  //치료비합계
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            string cureFrDt = dr["CureFrDt"] + "";
                            if (cureFrDt.CompareTo(strAcdtDt) >= 0)
                            {
                                db4DgnsNm += (db4DgnsNm != "" ? "\n" : "") + dr["DgnsNm"];
                                db4VstHosp += (db4VstHosp != "" ? "\n" : "") + dr["VstHosp"];
                                db4CureFrDt += (db4CureFrDt != "" ? "\n" : "") + Utils.DateFormat(dr["CureFrDt"], "yyyy.MM.dd") + "~" + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
                                string tmp = "";
                                tmp += (tmp != "" ? "\n" : "");
                                if (dr["CureFrDt"] + "" != "") tmp += Utils.DateFormat(dr["CureFrDt"], "yyyy.MM.dd");
                                if (dr["CureToDt"] + "" != "") tmp += "~" + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
                                if (dr["VstHospResn"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["VstHospResn"];
                                if (dr["TestNmRslt"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["TestNmRslt"];
                                if (dr["DoctDgns"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["DoctDgns"];
                                if (dr["CureCnts"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["CureCnts"];
                                if (dr["CureMjrCnts"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["CureMjrCnts"];
                                db4CureCnts += (db4CureCnts != "" ? "\n" : "") + tmp;
                            }
                            else
                            {
                                string tmp = "";
                                tmp += (tmp != "" ? "\n" : "");
                                if (dr["CureFrDt"] + "" != "") tmp += Utils.DateFormat(dr["CureFrDt"], "yyyy.MM.dd");
                                if (dr["CureToDt"] + "" != "") tmp += "~" + Utils.DateFormat(dr["CureToDt"], "yyyy.MM.dd");
                                if (dr["TestNmRslt"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["TestNmRslt"];
                                if (dr["MedHstr"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["MedHstr"];
                                if (dr["CureCnts"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["CureCnts"];
                                if (dr["CureMjrCnts"] + "" != "") tmp += (tmp != "" ? "\n" : "") + dr["CureMjrCnts"];
                                db4MedHstr += (db4MedHstr != "" ? "\n" : "") + tmp;
                            }
                            //치료비
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "CureFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "CureToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "Medfee")
                                {
                                    db4Medfee += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);
                            }
                        }
                        rUtil.ReplaceTables(lstTable, "@db4DgnsNm@", db4DgnsNm);
                        rUtil.ReplaceTables(lstTable, "@db4VstHosp@", db4VstHosp);
                        rUtil.ReplaceTables(lstTable, "@db4CureFrDt@", db4CureFrDt);
                        rUtil.ReplaceTables(lstTable, "@db4CureCnts@", db4CureCnts);
                        rUtil.ReplaceTables(lstTable, "@db4MedHstr@", db4MedHstr);
                        rUtil.ReplaceTableRow(oTableC.GetRow(dtB.Rows.Count + 1), "@db4Medfee@", Utils.AddComma(db4Medfee));
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTableE != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTableE.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        string tmp = "";
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"];
                        }
                        rUtil.ReplaceTextAllParagraph(doc, "@B6InsurObjDvs@", tmp);
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (oTableA != null)
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
                                    rUtil.ReplaceTableRow(oTableA.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        string tmp = "";
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            tmp += (tmp != "" ? "\n" : "");
                            tmp += dr["OthInsurCo"];
                            if (dr["OthInsurPrdt"] + "" != "") tmp += " " + dr["OthInsurPrdt"];
                            if (dr["OthCltrSpcCtrt"] + "" != "") tmp += " " + dr["OthCltrSpcCtrt"];
                            if (dr["OthCtrtDt"] + "" != "") tmp += " " + Utils.DateFormat(dr["OthCtrtDt"], "yyyy.MM.dd");
                            if (dr["OthCtrtExprDt"] + "" != "") tmp += "~" + Utils.DateFormat(dr["OthCtrtExprDt"], "yyyy.MM.dd");
                            if (dr["OthInsurRegsAmt"] + "" != "") tmp += ", " + Utils.AddComma(dr["OthInsurRegsAmt"]);

                        }
                        rUtil.ReplaceTextAllParagraph(doc, "@db8OthInsur@", tmp);
                        rUtil.ReplaceTables(lstTable, "@db8OthInsur@", tmp);
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTableF != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = i * 2;
                                int rmdr = 0;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTableF.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 50000L, 50000L, 4000000L, 3000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTableF.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTableG != null)
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
                                    rUtil.ReplaceTableRow(oTableG.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (oTableD != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTableD.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    string db13ExpsCmnt1 = "";
                    string db13ExpsCmnt2 = "";
                    string db13ExpsCmnt3 = "";
                    string db13ExpsCmnt4 = "";
                    string db13ExpsCmnt5 = "";
                    string db13ExpsCmnt6 = "";
                    string db13ExpsCmnt7 = "";
                    string db13ExpsCmnt8 = "";
                    string db13ExpsCmnt9 = "";
                    double db13ExpsLosAmt4 = 0;   //향후치료비
                    double db13ExpsLosAmt7 = 0;   //과실상계
                    double db13ExpsLosAmt8 = 0;   //위자료

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
                    if (dtB != null)
                    {
                        //1.치료비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        double dReq = 0;
                        double dAmt = 0;
                        string sExpsCmnt = "";
                        string sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt1 += (db13ExpsCmnt1 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        TableRow oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt1@", sExpsCmnt);
                        }

                        //2.휴업손해
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt2 += (db13ExpsCmnt2 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt2@", sExpsCmnt);
                        }

                        //3.상실수익
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt3 += (db13ExpsCmnt3 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt3@", sExpsCmnt);
                        }

                        //4.향후치료비
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        TableRow oRowBase = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt4@");
                        dReq = 0;
                        dAmt = 0;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableH, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsSubHed4@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt4@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt4@", drs[i]["ExpsCmnt"] + "");
                            }
                            db13ExpsCmnt4 += (db13ExpsCmnt4 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        db13ExpsLosAmt4 = dAmt;

                        //5.개호비
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt5 += (db13ExpsCmnt5 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt5@", sExpsCmnt);
                        }

                        //6.기타손해
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        oRowBase = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt6@");
                        dReq = 0;
                        dAmt = 0;
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (i == drs.Length - 1) oRow = oRowBase;
                            else oRow = rUtil.TableInsertBeforeRow(oTableH, oRowBase);
                            if (oRow != null)
                            {
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsSubHed6@", drs[i]["ExpsSubHed"] + "");
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt6@", Utils.AddComma(drs[i]["ExpsLosAmt"] + ""));
                                rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt6@", drs[i]["ExpsCmnt"] + "");
                            }
                            db13ExpsCmnt6 += (db13ExpsCmnt6 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }

                        //91.소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt91@", Utils.AddComma(dAmt));
                        }

                        //7.과실부담금
                        drs = dtB?.Select("ExpsGrp = 7");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt7 += (db13ExpsCmnt7 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt7@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt7@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt7@", sExpsCmnt);
                        }
                        db13ExpsLosAmt7 = dAmt;

                        //8.위자료
                        drs = dtB?.Select("ExpsGrp = 8");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt8 += (db13ExpsCmnt8 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt8@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt8@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt8@", sExpsCmnt);
                        }
                        db13ExpsLosAmt8 = dAmt;

                        //9.자기부담금
                        drs = dtB?.Select("ExpsGrp = 9");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                            db13ExpsCmnt9 += (db13ExpsCmnt9 != "" ? "\n" : "") + drs[i]["ExpsCmnt"];
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt9@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt9@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsCmnt9@", sExpsCmnt);
                        }

                        //92.합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sExpsCmnt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsLosReq"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsLosAmt"] + "");
                            if (Utils.ToInt(drs[i]["ExpsSeq"]) == 1)
                            {
                                sExpsCmnt = drs[i]["ExpsCmnt"] + "";
                                sExpsBss = drs[i]["ExpsBss"] + "";
                            }
                        }
                        oRow = rUtil.GetTableRow(oTableH?.Elements<TableRow>(), "@B13ExpsLosAmt92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B13ExpsLosAmt92@", Utils.AddComma(dAmt));
                        }
                    }
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt1@", db13ExpsCmnt1);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt2@", db13ExpsCmnt2);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt3@", db13ExpsCmnt3);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt4@", db13ExpsCmnt4);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt5@", db13ExpsCmnt5);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt6@", db13ExpsCmnt6);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt7@", db13ExpsCmnt7);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt8@", db13ExpsCmnt8);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsCmnt9@", db13ExpsCmnt9);
                    rUtil.ReplaceTables(lstTable, "@db13ExpsLosAmt4@", Utils.AddComma(db13ExpsLosAmt4));
                    rUtil.ReplaceTables(lstTable, "@db13ExpsLosAmt7@", Utils.AddComma(db13ExpsLosAmt7));
                    rUtil.ReplaceTables(lstTable, "@db13ExpsLosAmt8@", Utils.AddComma(db13ExpsLosAmt8));

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
