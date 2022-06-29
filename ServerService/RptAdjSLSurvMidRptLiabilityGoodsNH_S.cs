﻿using System;
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
    public class RptAdjSLSurvMidRptLiabilityGoodsNH_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvMidRptLiabilityGoodsNH_S(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2576_서식_농협_진행보고서(배책-차량, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2576_서식_농협_진행보고서(배책-차량, 간편).docx";
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
                rptName = "진행보고서_배책_차량, 간편(" + sfilename + ").docx";
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
                    Table oTableC = rUtil.GetTable(lstTable, "@B17ExpsDoLosAmt1@");

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        //기타비용 제목
                        if (!dtB.Columns.Contains("DoOthExpsHedT")) dtB.Columns.Add("DoOthExpsHedT");
                        {
                            if (Utils.ConvertToString(dr["DoOthExpsHed"]) == "")
                            {
                                dr["DoOthExpsHedT"] = "기타 비용";
                            }
                            else
                            {
                                dr["DoOthExpsHedT"] = dr["DoOthExpsHed"];
                            }
                        }

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
                            if (col.ColumnName == "CclsExptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");

                            //4. 추정지급보험금 표
                            if (col.ColumnName == "DoFixAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurAmt") sValue = Utils.AddComma(sValue);
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
                            if (col.ColumnName == "LeadAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "ChrgAdjLicSerl")
                            {
                                if (sValue != "") sValue = "손해사정 등록번호 : 제 " + sValue + " 호";
                            }
                            if (col.ColumnName == "BistLicSerl")
                            {
                                if (sValue != "") sValue = "보 조 인 등록번호 : 제 " + sValue + " 호";
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
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
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

                        tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"] + "/" + dr["ObjStrtRmk"];
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

                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
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



                    //세부 평가 내역
                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
                    if (dtB != null)
                    {
                        //1.수리비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        double dAmt = 0;
                        string sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";

                        }
                        TableRow oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt1@", sEvatRslt);
                        }

                        //2.휴차료
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt2@", sEvatRslt);
                        }

                        //3.대차료
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt3@", sEvatRslt);
                        }

                        //4.기타비용
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt4@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt4@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt4@", sEvatRslt);
                        }

                        //*소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt91@", Utils.AddComma(dAmt));
                        }

                        //5.과실부담금
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt5@", sEvatRslt);
                        }

                        //*합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt92@", Utils.AddComma(dAmt));
                        }

                        //6.자기부담금
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock17"].Rows.Add() };
                        dAmt = 0;
                        sEvatRslt = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B17ExpsDoLosAmt6@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B17ExpsDoLosAmt6@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B17EvatRslt6@", sEvatRslt);
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
