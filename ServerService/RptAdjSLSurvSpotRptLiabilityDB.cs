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
    public class RptAdjSLSurvSpotRptLiabilityDB
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvSpotRptLiabilityDB(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisSpotRprtLiabilityPrintDB";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2586_서식_DB_현장보고서(배책).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_2586_서식_DB_현장보고서(배책).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                DataTable dtBT = pds.Tables["DataBlock17"];
                if (dtBT != null && dtBT.Rows.Count > 0)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2586_서식_DB_현장보고서(배책)_Pict.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    sRet = SetSample1Pict(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }
                sSampleDocx = myPath + @"\보고서\출력설계_2586_서식_DB_현장보고서(배책)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                sRet = SetSample1Tail(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);

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
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "현장보고서_배책DB(" + sfilename + ").docx";
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
                    Table oTbl피해현황 = rUtil.GetTable(lstTable, "@B11DmobStrtCnts@"); // 4.피해현황
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B13RmnObjCost@"); //7.잔존물 표1
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B13SucBidDt@"); //7.잔존물 표2
                    Table oTbl현장배치도 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@"); //9.사고상황 - 가.현장배치도

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (oTbl피해현황 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl피해현황, 1, dtB.Rows.Count - 1);
                        }
                    }

                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl당사, 1, drs.Length - 1);
                        }
                    }
                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    if (dtB != null)
                    {
                        //9.현장배치도
                        if (oTbl현장배치도 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl현장배치도, 0, 2, dtB.Rows.Count - 1);
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
                    Table oTbl손해내용 = rUtil.GetTable(lstTable, "@B11InsurRegsAmt@"); // 3.손해내용
                    Table oTbl피해현황 = rUtil.GetTable(lstTable, "@B11DmobStrtCnts@"); // 4.피해현황
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B13RmnObjCost@"); //7.잔존물 표1
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B13SucBidDt@"); //7.잔존물 표2
                    Table oTbl현장배치도 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@"); //9.사고상황 - 가.현장배치도

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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "CclsExptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "VitmTel") sValue = (sValue == "" ? "-" : sValue);
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

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);//보상한도액
                            if (col.ColumnName == "EstmLosAmt") sValue = Utils.AddComma(sValue);//추정손해액
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);//공제금액
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTable(oTbl손해내용, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
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
                                rUtil.ReplaceTableRow(oTbl피해현황.GetRow(i + 1), sKey, sValue);
                            }
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    var db6OthInsur = "";

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);

                                if (col.ColumnName == "OthInsurCo")
                                {
                                    var OthInsurCo = dr["OthInsurCo"] + "";
                                    var OthInsurPrdt = dr["OthInsurPrdt"] + "";
                                    db6OthInsur += OthInsurCo + " " + OthInsurPrdt + "\n";
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }

                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 1");
                    sPrefix = "B13";
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
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

                    drs = pds.Tables["DataBlock13"]?.Select("TrtCd % 10 = 2");
                    sPrefix = "B13";
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
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

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl현장배치도 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl현장배치도.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl현장배치도.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    var db12InsurObjDvs = "";
                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //보험목적물
                                if (col.ColumnName == "ObjSymb")
                                {
                                    var InsurObjDvs = dr["InsurObjDvs"] + "";
                                    db12InsurObjDvs += InsurObjDvs + "\n";
                                }

                                //담보위험
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "1")
                                    {
                                        sValue = "보험 가입";
                                    }
                                    else
                                    {
                                        sValue = "보험 미가입";
                                    }
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }

                    var db14InsurObjNm = "";
                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //미가입목적물
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "0" || sValue == "")
                                    {
                                        var InsurObjNm = dr["InsurObjNm"] + "";
                                        db14InsurObjNm += InsurObjNm + "\n";
                                    }

                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTables(lstTable, sKey, sValue);
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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B17AcdtPictImage@");

                    dtB = pds.Tables["DataBlock17"];
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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B17AcdtPictImage@");

                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
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
                                int rmdr = i % 2 + 1;

                                TableRow xrow1 = oTbl현장사진.GetRow(rnum);

                                sKey = rUtil.GetFieldName(sPrefix, "ObjNm");
                                sValue = dr["ObjNm"] + "";
                                rUtil.SetText(xrow1.GetCell(0), sKey, sValue);

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 3000000L, 2400000L);
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

        private string SetSample1Tail(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    Table oTbl사고현장사진 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@"); //3.사진
                    Table oTbl사고관련자연락처 = rUtil.GetTable(lstTable, "@B8VitmNm@"); //사고관련자연락처

                    dtB = pds.Tables["DataBlock15"];
                    if (dtB != null)
                    {
                        //3.사고현장사진
                        if (oTbl사고현장사진 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl사고현장사진, 0, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    if (dtB != null)
                    {
                        if (oTbl사고관련자연락처 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl사고관련자연락처, 3, 1, dtB.Rows.Count - 1);
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
                    Table oTbl사고현장사진 = rUtil.GetTable(lstTable, "@B15AcdtPictImage@"); //3.사고현장사진
                    Table oTbl사고관련자연락처 = rUtil.GetTable(lstTable, "@B8VitmNm@"); //사고관련자연락처
                    Table oTbl체크리스트 = rUtil.GetTable(lstTable, "@db12InsurObjDvs@"); //체크리스트

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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "CclsExptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcptDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "MonSellAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "VitmTel") sValue = (sValue == "" ? "-" : sValue);
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

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);//보상한도액
                            if (col.ColumnName == "EstmLosAmt") sValue = Utils.AddComma(sValue);//추정손해액
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);//공제금액
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    var db6OthInsur = "";

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);

                                if (col.ColumnName == "OthInsurCo")
                                {
                                    var OthInsurCo = dr["OthInsurCo"] + "";
                                    var OthInsurPrdt = dr["OthInsurPrdt"] + "";
                                    db6OthInsur += OthInsurCo + " " + OthInsurPrdt + "\n";
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db6OthInsur@", db6OthInsur);


                    dtB = pds.Tables["DataBlock15"];
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (oTbl사고현장사진 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                sValue = dr["AcdtPictImage"] + "";
                                TableRow xrow1 = oTbl사고현장사진.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl사고현장사진.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
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
                                if (col.ColumnName == "VitmTel") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                                if (col.ColumnName == "VitmChrgTel") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                                rUtil.ReplaceTableRow(oTbl사고관련자연락처.GetRow(i + 3), sKey, sValue);
                            }
                        }
                    }

                    var db12InsurObjDvs = "";
                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;

                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //보험목적물
                                if (col.ColumnName == "ObjSymb")
                                {
                                    var InsurObjDvs = dr["InsurObjDvs"] + "";
                                    db12InsurObjDvs += InsurObjDvs + "\n";
                                }

                                //담보위험
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "1")
                                    {
                                        sValue = "보험 가입";
                                    }
                                    else
                                    {
                                        sValue = "보험 미가입";
                                    }
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db12InsurObjDvs@", db12InsurObjDvs);

                    var db14InsurObjNm = "";
                    dtB = pds.Tables["DataBlock14"];
                    sPrefix = "B14";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();

                        foreach (DataRow row in dtB.Rows)
                        {
                            DataRow dr = row;
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                //미가입목적물
                                if (col.ColumnName == "ObjInsurRegsFg")
                                {
                                    if (sValue == "0" || sValue == "")
                                    {
                                        var InsurObjNm = dr["InsurObjNm"] + "";
                                        db14InsurObjNm += InsurObjNm + "\n";
                                    }

                                    rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                    rUtil.ReplaceTables(lstTable, sKey, sValue);
                                }
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl체크리스트, "@db14InsurObjNm@", db14InsurObjNm);

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
