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
    public class RptAdjSLSurvRptLiabilityGoodsNH
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityGoodsNH(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2562_서식_농협_종결보고서(재물-대물, 배책-차량).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2562_서식_농협_종결보고서(재물-대물, 배책-차량).docx";
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
                rptName = "종결보고서_배책_대물(" + sfilename + ").docx";
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
                        Table oTbl총괄표 = rUtil.GetTable(lstTable, "@db6ObjInsurRegsAmt@");
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl총괄표, 1, dtB.Rows.Count - 1);
                        }
                        Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@db8OthInsur@");
                        TableRow oTblZ1Row = rUtil.GetTableRow(oTbl보험계약사항?.Elements<TableRow>(), "@B6InsurObjDvs@");
                        Table oTbl보상한도 = oTblZ1Row?.GetCell(1).Elements<Table>().FirstOrDefault();
                        if (oTbl보상한도 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableInsertRow(oTbl보상한도, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                        TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B7InsurGivObj@");
                        Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();
                        if (oTbl보험금지급처 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl보험금지급처, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock2"];
                    if (dtB != null)
                    {
                        Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B2AcdtPictImage@");
                        if (oTbl손해상황 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                rUtil.TableAddRow(oTbl손해상황, 2, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    if (dtB != null)
                    {
                        Table oTbl관련자연락처 = rUtil.GetTable(lstTable, "@B11AcdtPrsCcndGrp@");
                        if (oTbl관련자연락처 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl관련자연락처, 1, dtB.Rows.Count - 1);
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

                    drs = pds.Tables["DataBlock9"]?.Select("AcdtPictFg % 10 = 1 OR AcdtPictFg % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        Table oTbl첨부도면 = rUtil.GetTable(lstTable, "@B912AcdtPictImage@");
                        if (oTbl첨부도면 != null)
                        {
                            //테이블의 끝에 추가
                            for (int i = 1; i < drs.Length; i++)
                            {
                                rUtil.TableAddRow(oTbl첨부도면, 0, 1);
                                rUtil.TableAddRow(oTbl첨부도면, 1, 1);
                            }
                        }
                    }
                    drs = pds.Tables["DataBlock9"]?.Select("AcdtPictFg % 10 = 3 OR AcdtPictFg % 10 = 4 OR AcdtPictFg % 10 = 5");
                    if (drs != null && drs.Length > 0)
                    {
                        Table oTbl첨부사진 = rUtil.GetTable(lstTable, "@B9345AcdtPictImage@");
                        if (oTbl첨부사진 != null)
                        {
                            //테이블의 끝에 추가
                            for (int i = 1; i < drs.Length; i++)
                            {
                                rUtil.TableAddRow(oTbl첨부사진, 0, 1);
                                rUtil.TableAddRow(oTbl첨부사진, 1, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");
                        if (oTbl처리과정 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl처리과정, 1, dtB.Rows.Count - 1);
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
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@db6ObjInsurRegsAmt@");

                    Table oTbl보험계약사항 = rUtil.GetTable(lstTable, "@db8OthInsur@");
                    TableRow oTblZ1Row = rUtil.GetTableRow(oTbl보험계약사항?.Elements<TableRow>(), "@B6InsurObjDvs@");
                    Table oTbl보상한도 = oTblZ1Row?.GetCell(1).Elements<Table>().FirstOrDefault();

                    Table oTblA = rUtil.GetTable(lstTable, "@B1GivInsurCalcBrdn@");
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), "@B7InsurGivObj@");
                    Table oTbl보험금지급처 = oTblARow?.GetCell(0).Elements<Table>().FirstOrDefault();

                    Table oTbl피보험자관련사항 = rUtil.GetTable(lstTable, "피보험자 관련사항");

                    Table oTbl피해자관련사항 = rUtil.GetTable(lstTable, "피해자 관련사항");

                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B2AcdtPictImage@");

                    Table oTbl관련자연락처 = rUtil.GetTable(lstTable, "@B11AcdtPrsCcndGrp@");

                    Table oTbl별첨 = rUtil.GetTable(lstTable, "@B5FileNo@");

                    Table oTbl첨부도면 = rUtil.GetTable(lstTable, "@B912AcdtPictImage@");

                    Table oTbl첨부사진 = rUtil.GetTable(lstTable, "@B9345AcdtPictImage@");

                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");

                    var db1SurvAsgnEmpManRegNo = ""; //조사자 손해사정등록번호
                    var db1SurvAsgnEmpAssRegNo = ""; //조사자 보조인 등록번호
                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (dr["IsrdMjrPrdt"] + "" == "")
                        {
                            TableRow oTblCARow = rUtil.GetTableRow(oTbl피보험자관련사항?.Elements<TableRow>(), "@B1IsrdMjrPrdt@");
                            rUtil.TableRemoveRow(oTbl피보험자관련사항, oTblCARow);
                        }
                        if (dr["MonSellAmt"] + "" == "")
                        {
                            TableRow oTblCARow = rUtil.GetTableRow(oTbl피보험자관련사항?.Elements<TableRow>(), "@B1MonSellAmt@");
                            rUtil.TableRemoveRow(oTbl피보험자관련사항, oTblCARow);
                        }
                        if (dr["IsrdChrg"] + "" == "")
                        {
                            TableRow oTblCARow = rUtil.GetTableRow(oTbl피보험자관련사항?.Elements<TableRow>(), "@B1IsrdChrg@");
                            rUtil.TableRemoveRow(oTbl피보험자관련사항, oTblCARow);
                        }

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
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "IsrdOpenDt")
                            {
                                sValue = (sValue == "" ? "" : Utils.Mid(sValue, 1, 4) + "년 " + Utils.Mid(sValue, 5, 6) + "월");
                            }
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeConv(sValue, ":", "SHORT");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
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
                        if (oTbl손해상황 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            if (dtB.Rows.Count % 2 == 1) dtB.Rows.Add();  //두번째 칸을 클리어 해주기 위해서 추가
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                                int rmdr = i % 2 + 1;

                                TableRow xrow1 = oTbl손해상황.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(0), "@B2ObjNm@", dr["ObjNm"] + "");
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
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
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTable(oTbl피해자관련사항, sKey, sValue);
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
                        double db6ObjInsurRegsAmt = 0;
                        double db6ObjTotAmt = 0;
                        double db6ObjSelfBearAmt = 0;
                        double db6ObjGivInsurAmt = 0;
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            tmp += (tmp != "" ? "\n" : "") + dr["InsurObjDvs"];
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "ObjInsurRegsAmt")
                                {
                                    db6ObjInsurRegsAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
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
                                if (col.ColumnName == "LosAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "NglgSetoffAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "EvatStdLosCnts")
                                {
                                    // (값이 있을 경우 한 행을 띄우고 출력)
                                    if (sValue != "") sValue = "\n" + sValue;
                                }
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl보상한도.GetRow(i + 1), sKey, sValue);
                                sKey = rUtil.GetFieldName("B61", col.ColumnName);
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }

                        }
                        rUtil.ReplaceTextAllParagraph(doc, "@db6InsurObjDvs@", tmp);
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjInsurRegsAmt@", Utils.AddComma(db6ObjInsurRegsAmt));
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjTotAmt@", Utils.AddComma(db6ObjTotAmt));
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjSelfBearAmt@", Utils.AddComma(db6ObjSelfBearAmt));
                        rUtil.ReplaceTableRow(oTbl총괄표.GetRow(dtB.Rows.Count + 1), "@db6ObjGivInsurAmt@", Utils.AddComma(db6ObjGivInsurAmt));
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

                    drs = pds.Tables["DataBlock9"]?.Select("AcdtPictFg % 10 = 1 OR AcdtPictFg % 10 = 2");
                    if (drs != null)
                    {
                        if (oTbl첨부도면 != null)
                        {
                            if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock9"].Rows.Add() };
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                int rnum = i * 2;
                                int rmdr = 0;

                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl첨부도면.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), "@B912AcdtPictImage@", "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 5700000L, 4200000L);
                                }
                                catch { }

                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl첨부도면.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), "@B912AcdtPictCnts@", sValue);
                            }
                        }
                    }

                    drs = pds.Tables["DataBlock9"]?.Select("AcdtPictFg % 10 = 3 OR AcdtPictFg % 10 = 4 OR AcdtPictFg % 10 = 5");
                    if (drs != null)
                    {
                        if (oTbl첨부사진 != null)
                        {
                            if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock9"].Rows.Add() };
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                int rnum = i * 2;
                                int rmdr = 0;

                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl첨부사진.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), "@B9345AcdtPictImage@", "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 4100000L, 3000000L);
                                }
                                catch { }

                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl첨부사진.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), "@B9345AcdtPictCnts@", sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl처리과정 != null)
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
                                    rUtil.ReplaceTableRow(oTbl처리과정.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (oTbl관련자연락처 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    rUtil.ReplaceTableRow(oTbl관련자연락처.GetRow(i + 1), sKey, sValue);
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
