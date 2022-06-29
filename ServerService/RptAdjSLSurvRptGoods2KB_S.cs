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
    public class RptAdjSLSurvRptGoods2KB_S
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptGoods2KB_S(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtGoodsPrintKB";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2592_서식_KB_종결보고서(재물-대물, 간편).xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSample1Docx = myPath + @"\보고서\출력설계_2592_서식_KB_종결보고서(재물-대물, 간편).docx";
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
                rptName = "종결보고서_재물-대물, 간편 - KB(" + sfilename + ").docx";
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
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl유첨서류 = rUtil.GetTable(lstTable, "@B9FileNo@");
                    Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");
                    Table oTbl보험금지급처 = rUtil.GetTable(lstTable, "@B16InsurGivObj@");
                    Table oTbl피해물및손해사정 = rUtil.GetTable(lstTable, "@B17ObjCstrSpcf@");
                    TableRow oTblR복구수리비용 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@B17ObjCstrSpcf@");

                    
                    //TableCell oTblC = oTblR복구수리비용?.GetCell(1);
                    //Table oTableC복구수리비용 = oTblR복구수리비용?.GetCell(1).Elements<Table>(). FirstOrDefault();

                    
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl당사, 1, drs.Length - 1);
                        }
                    }
                    else oTbl당사.Remove();
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }
                    else oTbl옥션.Remove();
                    
                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl유첨서류 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl유첨서류, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl진행사항 != null)
                        {
                            rUtil.TableAddRow(oTbl진행사항, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock16"];
                    if (dtB != null)
                    {
                        if (oTbl보험금지급처 != null)
                        {
                            //테이블의 중간에 추가
                            rUtil.TableAddRow(oTbl보험금지급처, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
                    if (dtB != null)
                    {
                        if (oTbl피해물및손해사정 != null)
                        {
                            rUtil.TableInsertRow(oTbl피해물및손해사정, 4, dtB.Rows.Count - 1);
                            rUtil.TableMergeCells(oTbl피해물및손해사정, 0, 0, 4, dtB.Rows.Count + 4);
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
                    Table oTbl피해물및손해사정 = rUtil.GetTable(lstTable, "@db3InsurObjDvsText@");
                    TableRow oTblR복구수리비용합계 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@db17ObjCstrAmt@");
                    TableRow oTblR손해배상금 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@db18LosEvatTot@");
                    TableRow oTblR과실상계 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@db18NglgSetoffAmt@");
                    TableRow oTblR자기부담금 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@db18ObjSelfBearAmt@");
                    TableRow oTblR지급보험금 = rUtil.GetTableRow(oTbl피해물및손해사정?.Elements<TableRow>(), "@db18ObjGivInsurAmt@");
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl유첨서류 = rUtil.GetTable(lstTable, "@B9FileNo@");
                    Table oTbl진행사항 = rUtil.GetTable(lstTable, "@B10PrgMgtDt@");
                    Table oTbl보험금지급처 = rUtil.GetTable(lstTable, "@B16InsurGivObj@");
                    Table oTbl빈테이블 = rUtil.GetTable(lstTable, "@Table@");
                    
                    

                    //sKey = "@B1AcdtCnts@";
                    //Table oTblA = rUtil.GetTable(lstTable, sKey);
                    //sKey = "@B2AcdtPictImage1@";
                    //TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    //Table oTableA = oTblARow?.GetCell(1).Elements<Table>().FirstOrDefault();

                    ////sKey = "@B1AcdtCnts@";
                    ////Table oTblB = rUtil.GetTable(lstTable, sKey);
                    ////sKey = "@B2AcdtPictImage2@";
                    ////TableRow oTblBRow = rUtil.GetTableRow(oTblB?.Elements<TableRow>(), sKey);
                    ////Table oTableB = oTblBRow?.GetCell(1).Elements<Table>().FirstOrDefault();


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
                            if (col.ColumnName == "EmpEmail") sValue = (sValue == "" ? "" : sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue);
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
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd.");
                            if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }


                    var db3InsurObjDvsText = "";
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
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

                                if (col.ColumnName == "ObjSymb")
                                {
                                    var InsurObjDvs = dr["InsurObjDvs"] + "";
                                    if (!(InsurObjDvs == null) && !(InsurObjDvs == "")) { db3InsurObjDvsText += InsurObjDvs + "\n"; }
                                }
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl피해물및손해사정, "@db3InsurObjDvsText@", db3InsurObjDvsText); //4.피해물 및 손해사정 - 피해물



                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    sPrefix = "B8";
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

                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    sPrefix = "B8";
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
                                    if (col.ColumnName == "FileAmt") sValue = Utils.AddComma(sValue == "" || sValue == "0" ? "1" : sValue) + "부";
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
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

                    dtB = pds.Tables["DataBlock16"];
                    sPrefix = "B16";
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
                                    if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue); //지급보험금
                                    rUtil.ReplaceTableRow(oTbl보험금지급처.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    double db17ObjCstrAmt = 0;
                    dtB = pds.Tables["DataBlock17"];
                    sPrefix = "B17";
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
                                if (col.ColumnName == "ObjCstrAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "ObjCstrAmt") // 복구수리비용 합계
                                {
                                    db17ObjCstrAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl피해물및손해사정.GetRow(i + 4), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR복구수리비용합계, "@db17ObjCstrAmt@", Utils.AddComma(db17ObjCstrAmt));


                    double db18LosEvatTot = 0;
                    double db18NglgSetoffAmt = 0;
                    double db18ObjSelfBearAmt = 0;
                    double db18ObjGivInsurAmt = 0;
                    dtB = pds.Tables["DataBlock18"];
                    sPrefix = "B18";
                    if (dtB != null)
                    {
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            
                            foreach (DataColumn col in dtB.Columns) {
                            
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "LosEvatTot") // 손해배상금 합계
                            {
                                db18LosEvatTot += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "NglgSetoffAmt") // 과실상계 합계
                            {
                                db18NglgSetoffAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "ObjSelfBearAmt") // 자기부담금 합계
                            {
                                db18ObjSelfBearAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                            if (col.ColumnName == "ObjGivInsurAmt") // 지급보험금 합계
                            {
                                db18ObjGivInsurAmt += Utils.ToDouble(sValue);
                                sValue = Utils.AddComma(sValue);
                            }
                              
                            }
                        }
                        rUtil.ReplaceTableRow(oTblR손해배상금, "@db18LosEvatTot@", Utils.AddComma(db18LosEvatTot));
                        rUtil.ReplaceTableRow(oTblR과실상계, "@db18NglgSetoffAmt@", Utils.AddComma(db18NglgSetoffAmt));
                        rUtil.ReplaceTableRow(oTblR자기부담금, "@db18ObjSelfBearAmt@", Utils.AddComma(db18ObjSelfBearAmt));
                        rUtil.ReplaceTableRow(oTblR지급보험금, "@db18ObjGivInsurAmt@", Utils.AddComma(db18ObjGivInsurAmt));
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
