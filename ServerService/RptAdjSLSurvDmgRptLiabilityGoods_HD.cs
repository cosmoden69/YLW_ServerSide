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
    public class RptAdjSLSurvDmgRptLiabilityGoods_HD
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvDmgRptLiabilityGoods_HD(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintDmg";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2642_서식_교부용 손해사정서_배상_현대해상.xsd";
                //string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_2642_서식_교부용 손해사정서_배상_현대해상.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
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
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }

                rptName = "손해사정서_배책(" + sfilename + ").docx";
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

                    double db1InsurRegsAmt = 0;
                    double db1EstmLosAmt = 0;
                    double db1EstmLosAmt2 = 0;
                    double db1SelfBearAmt = 0;
                    double db1GivInsurAmt = 0;

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
                            if (col.ColumnName == "CltrStpltRspsbFg") continue;
                            if (col.ColumnName == "CltrStpltRspsbSrc") continue;
                            if (col.ColumnName == "CltrStpltRspsbBss") continue;
                            if (col.ColumnName == "RgtCpstOpni") continue;
                            if (col.ColumnName == "RgtCpstCnclsRmk") continue;
                            if (col.ColumnName == "RgtCpstSrc") continue;
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
                            if (col.ColumnName == "AcptRgstDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "CclsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "EmpCellPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            //보상한도액
                            if (col.ColumnName == "InsurRegsAmt")
                            {
                                db1InsurRegsAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "DmInsurRegsAmt")
                            {
                                db1InsurRegsAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            //손해액
                            if (col.ColumnName == "DiSubTotAmt")
                            {
                                db1EstmLosAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "DmEstmLosAmt")
                            {
                                db1EstmLosAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            //손해배상금
                            if (col.ColumnName == "DiTotAmt")
                            {
                                db1EstmLosAmt2 += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "DmEstmLosAmt2")
                            {
                                db1EstmLosAmt2 += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            //자기부담금
                            if (col.ColumnName == "DiSelfBearAmt")
                            {
                                db1SelfBearAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "DmSelfBearAmt")
                            {
                                db1SelfBearAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            //지급보험금
                            if (col.ColumnName == "DiGivInsurAmt")
                            {
                                db1GivInsurAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "DmGivInsurAmt")
                            {
                                db1GivInsurAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
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
                        rUtil.ReplaceTables(lstTable, "@db1InsurRegsAmt@", Utils.AddComma(db1InsurRegsAmt));
                        rUtil.ReplaceTables(lstTable, "@db1EstmLosAmt@", Utils.AddComma(db1EstmLosAmt));
                        rUtil.ReplaceTables(lstTable, "@db1EstmLosAmt2@", Utils.AddComma(db1EstmLosAmt2));
                        rUtil.ReplaceTables(lstTable, "@db1SelfBearAmt@", Utils.AddComma(db1SelfBearAmt));
                        rUtil.ReplaceTables(lstTable, "@db1GivInsurAmt@", Utils.AddComma(db1GivInsurAmt));

                        string db1Tmp = "";
                        sValue = Utils.ConvertToString(dr["CltrStpltRspsbFg"]);
                        if (sValue == "0") sValue = "면책";
                        else if (sValue == "1") sValue = "부책";
                        else sValue = "";
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        sValue = Utils.ConvertToString(dr["CltrStpltRspsbBss"]);
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        sValue = Utils.ConvertToString(dr["CltrStpltRspsbSrc"]);
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        rUtil.ReplaceTables(lstTable, "@db1CltrStpltRspsb@", db1Tmp);

                        db1Tmp = "";
                        sValue = Utils.ConvertToString(dr["RgtCpstOpni"]);
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        sValue = Utils.ConvertToString(dr["RgtCpstCnclsRmk"]);
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        sValue = Utils.ConvertToString(dr["RgtCpstSrc"]);
                        if (db1Tmp != "" && sValue != "") db1Tmp += "\n";
                        db1Tmp += sValue;
                        rUtil.ReplaceTables(lstTable, "@db1RgtCpst@", db1Tmp);
                    }

                    string db2DmgCnts = "";     //피해내용
                    string db2Vitm2 = "";       //피해자
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];

                            string sdmgTmp = Utils.ConvertToString(dr["DmgCnts"]);
                            string svitmTmp = Utils.ConvertToString(dr["VitmNm"]);
                            if (db2DmgCnts != "" && svitmTmp != "") db2DmgCnts += "\n";
                            db2DmgCnts += svitmTmp + " : " + sdmgTmp;

                            if (db2Vitm2 != "" && svitmTmp != "") db2Vitm2 += ",";
                            db2Vitm2 += svitmTmp;
                        }
                    }
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];

                            string sdmobTmp = Utils.ConvertToString(dr["DmobNm"]);
                            string sdmgTmp = Utils.ConvertToString(dr["DmobDmgStts"]);
                            if (db2DmgCnts != "" && sdmobTmp != "") db2DmgCnts += "\n";
                            db2DmgCnts += sdmobTmp + " : " + sdmgTmp;
                        }
                    }
                    rUtil.ReplaceTables(lstTable, "@db2DmgCnts@", db2DmgCnts);
                    rUtil.ReplaceTables(lstTable, "@db2Vitm2@", db2Vitm2);

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
