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
{//
    public class RptAdjSLRptSurvRptPers
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLRptSurvRptPers(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisCclsRprtPersPrint";
                security.methodId = "Print";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_1511_서식_종결보고서_표준.xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_1511_서식_종결보고서_표준.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                DataTable dtB = pds.Tables["DataBlock9"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_1511_서식_종결보고서_표준_Image.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    sRet = SetSample_Image(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
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
                dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "종결보고서_인보험(" + sfilename + ").docx";
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B3InsurPrdt@");
                    Table oTbl청구내용 = rUtil.GetTable(lstTable, "@B2AcdtDt@");
                    Table oTbl기본정보 = rUtil.GetTable(lstTable, "@B3InsurPrdt1@");
                    Table oTbl사고경위 = rUtil.GetTable(lstTable, "@B5ContentsDate@");
                    Table oTbl별첨자료 = rUtil.GetTable(lstTable, "@B8FileNo@");
                    Table oTbl지급조사수수료 = rUtil.GetTable(lstTable, "@B12InvcFee@");

                    dtB = pds.Tables["DataBlock2"];
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        string HasFeeTable = dtB.Rows[0]["HasFeeTable"] + "";
                        if (HasFeeTable != "1")
                        {
                            oTbl지급조사수수료.Remove();
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (oTbl계약사항 != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableAddRow(oTbl계약사항, 1, 1);
                                rUtil.TableAddRow(oTbl계약사항, 2, 1);
                            }
                        }
                        if (oTbl기본정보 != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableAddRow(oTbl기본정보, 1, 1);
                                rUtil.TableAddRow(oTbl기본정보, 2, 1);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl사고경위 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl사고경위, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (oTbl별첨자료 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl별첨자료, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableAddRow(oTable, 0, 1);
                                rUtil.TableAddRow(oTable, 1, 1);
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
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }


                    //dtB = pds.Tables["DataBlock5"];
                    //sPrefix = "B5";
                    //if (dtB != null)
                    //{
                    //    sKey = rUtil.GetFieldName(sPrefix, "ContentsDate");
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B3InsurPrdt@");
                    Table oTbl별첨자료 = rUtil.GetTable(lstTable, "@B8FileNo@");
                    Table oTbl기본정보 = rUtil.GetTable(lstTable, "@B3InsurPrdt1@");
                    Table oTbl사고경위 = rUtil.GetTable(lstTable, "@B5ContentsDate@");

                    Table oTbl손해사정내용 = rUtil.GetTable(lstTable, "@db13SurvCnts1@");
                    TableRow oTblR청구내용 = rUtil.GetTableRow(oTbl손해사정내용?.Elements<TableRow>(), "@db13SurvCnts1@");
                    TableRow oTblR확인내용 = rUtil.GetTableRow(oTbl손해사정내용?.Elements<TableRow>(), "@db13SurvCnts2@");
                    Table oTbl손해사정의견 = rUtil.GetTable(lstTable, "@db13SurvCnts3@");
                    Table oTbl약관규정및관련법규 = rUtil.GetTable(lstTable, "@db13SurvCnts5@");
                    Table oTbl조사결과 = rUtil.GetTable(lstTable, "@db13SurvCnts6@");
                    TableRow oTblR확인사항 = rUtil.GetTableRow(oTbl조사결과?.Elements<TableRow>(), "@db13SurvCnts6@");
                    TableRow oTblR결론 = rUtil.GetTableRow(oTbl조사결과?.Elements<TableRow>(), "@db13SurvCnts9@");
                    Table oTbl세부조사내용 = rUtil.GetTable(lstTable, "@db13SurvCnts11@");
                    TableRow oTblR피보험자면담사항 = rUtil.GetTableRow(oTbl세부조사내용?.Elements<TableRow>(), "@db13SurvCnts11@");
                    TableRow oTblR모집인면담사항 = rUtil.GetTableRow(oTbl세부조사내용?.Elements<TableRow>(), "@db13SurvCnts14@");
                    TableRow oTblR타보험사확인사항 = rUtil.GetTableRow(oTbl세부조사내용?.Elements<TableRow>(), "@db13SurvCnts15@");
                    TableRow oTblR민원예방활동 = rUtil.GetTableRow(oTbl세부조사내용?.Elements<TableRow>(), "@db13SurvCnts16@");
                    Table oTbl지급조사수수료 = rUtil.GetTable(lstTable, "@B12InvcFee@");

                    
                    var RaisResnFg = ""; //모집인 면담사항 행삭제 때문에 사용
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];

                        sKey = rUtil.GetFieldName(sPrefix, "AcptDt");
                        sValue = Utils.PadLeft(dr["AcptDtYear"], 4) + "년" + Utils.PadLeft(dr["AcptDtMonth"], 2) + "월 " + Utils.PadLeft(dr["AcptDtDays"], 2) + "일";
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        sKey = rUtil.GetFieldName(sPrefix, "LasRptDt");
                        sValue = Utils.PadLeft(dr["LasRptSbmsDtYear"], 4) + "년 " + Utils.PadLeft(dr["LasRptSbmsDtMonth"], 2) + "월 " + Utils.PadLeft(dr["LasRptSbmsDtDays"], 2) + "일";
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        foreach (DataColumn col in dtB.Columns)
                        {
                            //if (col.ColumnName == "AcptDtYear") continue;
                            //if (col.ColumnName == "AcptDtMonth") continue;
                            //if (col.ColumnName == "AcptDtDays") continue;
                            //if (col.ColumnName == "LasRptSbmsDtYear") continue;
                            //if (col.ColumnName == "LasRptSbmsDtMonth") continue;
                            //if (col.ColumnName == "LasRptSbmsDtDays") continue;
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "MjrCureFrDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "MjrCureToDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "SurvAsgnTeamLeadOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpHP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "RaisResnFg")
                            {
                                RaisResnFg = sValue; //모집인 면담사항 행삭제 때문에 사용
                                if (sValue == "Y")
                                {
                                    sValue = "징구";
                                }
                                else
                                {
                                    sValue = "징구하지 않음";
                                }
                            }
                            if (col.ColumnName == "GivObjRels")
                            {
                                if (sValue != "")
                                {
                                    sValue = "(" + sValue + ")";
                                }
                            }
                            if (col.ColumnName == "LeadAdjPhoto" || col.ColumnName == "ChrgAdjPhoto" || col.ColumnName == "SealPhotoLead" || col.ColumnName == "SealPhotoEmp")
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }


                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
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
                                if (col.ColumnName == "CtrtDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "InsurRegsAmt")
                                {
                                    if (sValue != "0" && sValue != "")
                                    {
                                        sValue = Utils.AddComma(sValue) + "원";
                                    }
                                }
                                rUtil.ReplaceTableRow(oTbl계약사항.GetRow((i * 2) + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl계약사항.GetRow((i * 2) + 2), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl기본정보.GetRow((i * 2) + 1), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl기본정보.GetRow((i * 2) + 2), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl사고경위 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "ContentsDate") sValue = Utils.DateConv(sValue, ".");
                                    rUtil.ReplaceTableRow(oTbl사고경위.GetRow(i + 1), sKey, sValue);
                                }
                                if (dr["ContentsType"] + "" == "1")
                                {
                                    rUtil.TableRowBackcolor(oTbl사고경위.GetRow(i + 1), "ABCDEF");
                                }
                            }
                        }
                    }
                    /*==============================================================================================================================================================*/
                    /*============================ 3.손해사정내용 =======================================================================================*/
                    //청구내용
                    var db13SurvCnts1 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181001");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts1 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR청구내용, "@db13SurvCnts1@", db13SurvCnts1);
                    
                    //확인내용
                    var db13SurvCnts2 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181002");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts2 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR확인내용, "@db13SurvCnts2@", db13SurvCnts2);

                    /*============================ 4.손해사정의견 =======================================================================================*/
                    // 1) 보험급지급여부
                    var db13SurvCnts3 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181003");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts3 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl손해사정의견, "@db13SurvCnts3@", db13SurvCnts3);
                    //rUtil.ReplaceTableRow(oTblR확인내용, "@db13SurvCnts2@", db13SurvCnts2);

                    // 2) 계약 전 알릴 의무 위반 여부
                    var db13SurvCnts4 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181004");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts4 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl손해사정의견, "@db13SurvCnts4@", db13SurvCnts4);

                    /*============================ 5.약관규정 및 관련 법규 =======================================================================================*/
                    // 5.약관규정 및 관련 법규
                    var db13SurvCnts5 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181005");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts5 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl약관규정및관련법규, "@db13SurvCnts5@", db13SurvCnts5);

                    /*============================ 3.조사결과 - 확인사항 =======================================================================================*/
                    // 1) 청구권 관련 청구병원 확인사항
                    var db13SurvCnts6 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181006");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts6 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR확인사항, "@db13SurvCnts6@", db13SurvCnts6);

                    // 2) 계약 전 알릴의무 관련 과거병력 확인사항
                    var db13SurvCnts7 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181007");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts7 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR확인사항, "@db13SurvCnts7@", db13SurvCnts7);

                    // 3) 후유장해 적정성 확인
                    var db13SurvCnts8 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181008");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts8 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR확인사항, "@db13SurvCnts8@", db13SurvCnts8);

                    /*============================ 3.조사결과 - 결론 =======================================================================================*/
                    // 1) 보험급 지급 여부
                    var db13SurvCnts9 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181009");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts9 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR결론, "@db13SurvCnts9@", db13SurvCnts9);

                    // 2) 계약 전 알릴 의무 위반 여부
                    var db13SurvCnts10 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181010");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts10 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR결론, "@db13SurvCnts10@", db13SurvCnts10);

                    /*============================ 4.세부 조사 내용 - 피보험자 면담사항 =======================================================================================*/
                    // 1) 직업 및 생활 환경
                    var db13SurvCnts11 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181011");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts11 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR피보험자면담사항, "@db13SurvCnts11@", db13SurvCnts11);

                    // 2) 보험 가입 경위
                    var db13SurvCnts12 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181012");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts12 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR피보험자면담사항, "@db13SurvCnts12@", db13SurvCnts12);

                    // 3) 가입 전 병력 및 주요사항 고지 여부
                    var db13SurvCnts13 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181013");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts13 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR피보험자면담사항, "@db13SurvCnts13@", db13SurvCnts13);

                    /*============================ 4.세부 조사 내용 - 모집인 면담사항 =======================================================================================*/
                    // 2) 모집경위서 주요 내용
                    var db13SurvCnts14 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181014");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts14 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR모집인면담사항, "@db13SurvCnts14@", db13SurvCnts14);
                    if (db13SurvCnts14 == "" && RaisResnFg == "") { oTblR모집인면담사항.Remove(); } //4.세부조사내용 - 모집인 면담사항 : (입력된 내용이 없을 경우, “모집인 면담사항” 행 전체를 삭제)

                    /*============================ 4.세부 조사 내용 - 타보험사 확인사항 =======================================================================================*/
                    // 4.세부 조사 내용 - 타보험사 확인사항
                    var db13SurvCnts15 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181015");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts15 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR타보험사확인사항, "@db13SurvCnts15@", db13SurvCnts15);

                    /*============================ 4.세부 조사 내용 - 민원예방활동 =======================================================================================*/
                    // 1) 조사과정 중 고객 불만사항
                    var db13SurvCnts16 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181016");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts16 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR민원예방활동, "@db13SurvCnts16@", db13SurvCnts16);

                    // 2) 불만사항에 대한 조치내용
                    var db13SurvCnts17 = "";
                    drs = pds.Tables["DataBlock13"]?.Select("SurvCntsCd = 300181017");
                    sPrefix = "B13";
                    if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock13"].Rows.Add() };
                    if (drs != null && drs.Length > 0)
                    {
                        for (int i = 0; i < drs.Length; i++)
                        {
                            DataRow dr = drs[i];
                            foreach (DataColumn col in dr.Table.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "SurvCnts") db13SurvCnts17 = sValue;
                                //rUtil.ReplaceTableRow(oTblR청구내용, sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTblR민원예방활동, "@db13SurvCnts17@", db13SurvCnts17);
                    /*==============================================================================================================================================================*/

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
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
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (i * 2) + 0;
                                int rmdr = 0;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                                sValue = dr["AcdtPictPath"] + "";
                                TableRow xrow1 = oTable.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 600000L, 50000L, 6000000L, 3500000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTable.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
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
                                    if (col.ColumnName == "PrgMgtDt") sValue = Utils.DateConv(sValue, ".");
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
                            }
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
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
                    if (dtB != null && oTbl지급조사수수료 != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            rUtil.ReplaceTable(oTbl지급조사수수료, sKey, sValue);
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

        private string SetSample_Image(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            for (int i = 1; i < dtB.Rows.Count; i++)
                            {
                                rUtil.TableAddRow(oTable, 0, 1);
                                rUtil.TableAddRow(oTable, 1, 1);
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

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (i * 2) + 0;
                                int rmdr = 0;

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictPath");
                                sValue = dr["AcdtPictPath"] + "";
                                TableRow xrow1 = oTable.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImage(xrow1.GetCell(rmdr), img, 600000L, 50000L, 6000000L, 3500000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTable.GetRow(rnum + 1);
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
