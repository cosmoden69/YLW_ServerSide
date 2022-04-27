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
    public class DlgAdjSLSurvRptLina
    {
        private string myPath = Application.StartupPath;

        public DlgAdjSLSurvRptLina(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisSurvRptLinaPrint";
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

                string sSampleXSD = myPath + @"\보고서\DlgAdjSLSurvRptLina.xsd";

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\DlgAdjSLSurvRptLina.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);

                //Console.WriteLine("{0} : {1}", DateTime.Now.ToString("HH:mm:ss"), "End");

                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string sfilename = "";
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "전문보고서_라이나(" + sfilename + ").docx";
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
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B13InsurName@");
                    Table oTbl조사위임 = rUtil.GetTable(lstTable, "@B2OurAcdtSurvDtlCodeNm@");
                    Table oTbl조사의뢰 = rUtil.GetTable(lstTable, "@B9FocusedIssue@");
                    Table oTbl조사결과 = rUtil.GetTable(lstTable, "@B9SurvResult@");
                    Table oTbl과거력 = rUtil.GetTable(lstTable, "@B6TreatFrTo@");
                    Table oTbl타사가입 = rUtil.GetTable(lstTable, "@B7OthInsurCompCodeNm@");
                    Table oTbl체크리스트 = rUtil.GetTable(lstTable, "@B4TermsRelayCnts@");
                    Table oTbl민원예방 = rUtil.GetTable(lstTable, "@B8CmplPrvtCodeNm@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B11PrgMgtDt@");

                    dtB = pds.Tables["DataBlock13"];
                    if (dtB != null)
                    {
                        //테이블의 끝에 추가
                        rUtil.TableInsertRow(oTbl계약사항, 1, dtB.Rows.Count - 1);
                    }

                    dtB = pds.Tables["DataBlock9"];
                    if (dtB != null)
                    {
                        if (oTbl조사의뢰 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl조사의뢰, 2, dtB.Rows.Count - 1);
                        }
                        if (oTbl조사결과 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableInsertRow(oTbl조사결과, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    if (dtB != null)
                    {
                        //테이블의 끝에 추가
                        rUtil.TableInsertRow(oTbl과거력, 1, dtB.Rows.Count - 1);
                    }

                    dtB = pds.Tables["DataBlock7"];
                    if (dtB != null)
                    {
                        //테이블의 중간에 삽입
                        rUtil.TableInsertRow(oTbl타사가입, 1, dtB.Rows.Count - 1);
                    }

                    dtB = pds.Tables["DataBlock8"];
                    if (dtB != null)
                    {
                        //테이블의 중간에 삽입
                        rUtil.TableInsertRow(oTbl민원예방, 1, dtB.Rows.Count - 1);
                    }

                    dtB = pds.Tables["DataBlock11"];
                    if (dtB != null)
                    {
                        //테이블의 중간에 삽입
                        rUtil.TableInsertRow(oTbl처리과정, 1, dtB.Rows.Count - 1);
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

                    IEnumerable<Table> lstTable = doc.Body.Elements<Table>();
                    Table oTbl계약사항 = rUtil.GetTable(lstTable, "@B13InsurName@");
                    Table oTbl조사위임 = rUtil.GetTable(lstTable, "@B2OurAcdtSurvDtlCodeNm@");
                    Table oTbl조사의뢰 = rUtil.GetTable(lstTable, "@B9FocusedIssue@");
                    Table oTbl조사결과 = rUtil.GetTable(lstTable, "@B9SurvResult@");
                    Table oTbl과거력 = rUtil.GetTable(lstTable, "@B6TreatFrTo@");
                    Table oTbl타사가입 = rUtil.GetTable(lstTable, "@B7OthInsurCompCodeNm@");
                    Table oTbl체크리스트 = rUtil.GetTable(lstTable, "@B4TermsRelayCnts@");
                    Table oTbl민원예방 = rUtil.GetTable(lstTable, "@B8CmplPrvtCodeNm@");
                    Table oTbl처리과정 = rUtil.GetTable(lstTable, "@B11PrgMgtDt@");

                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null && dtB.Rows.Count > 0)
                    {
                        DataRow dr = dtB.Rows[0];

                        sKey = rUtil.GetFieldName(sPrefix, "SurvReqDt");
                        sValue = Utils.DateFormat(dr["SurvReqDt"], "yyyy년 MM월 dd일");
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        sKey = rUtil.GetFieldName(sPrefix, "SurvCompDt");
                        sValue = Utils.DateFormat(dr["SurvCompDt"], "yyyy년 MM월 dd일");
                        rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "IsrdRegno1")
                            {
                                sKey = "@B1IsrdRegno@";
                                sValue = dr["IsrdRegno1"] + "-" + dr["IsrdRegno2"];
                            }
                            if (col.ColumnName == "SurvAcptDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "SurvReqDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "SurvCompDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "DelayRprtDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateConv(sValue, ".");
                            if (col.ColumnName == "SurvAsgnTeamLeadOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpOP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "SurvAsgnEmpHP") sValue = Utils.TelNumber(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjPhoto" || col.ColumnName == "SealPhotoLead" || col.ColumnName == "SealPhotoEmp")
                            {
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.ReplaceInternalImage(sKey, img);
                                }
                                catch { }
                                continue;
                            }
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock13"];
                    sPrefix = "B13";
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
                                if (col.ColumnName == "ContractDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "EffectiveDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "resurrectionDt") sValue = Utils.DateConv(sValue, ".");
                                rUtil.ReplaceTableRow(oTbl계약사항.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
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
                                rUtil.ReplaceTableRow(oTbl조사위임.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
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
                                if (col.ColumnName == "RsnOccurDt") sValue = Utils.DateConv(sValue, ".");
                                if (col.ColumnName == "ClaimReqAmt")
                                {
                                    if (sValue != "0" && sValue != "") sValue = Utils.AddComma(sValue) + "원";
                                }
                                if (col.ColumnName == "ClaimPayAmt")
                                {
                                    if (sValue != "0" && sValue != "") sValue = Utils.AddComma(sValue) + "원";
                                }
                                rUtil.ReplaceTableRow(oTbl조사의뢰.GetRow(i + 2), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl조사결과.GetRow(i + 1), sKey, sValue);
                            }
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
                                rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
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
                                if (col.ColumnName == "TreatFrDt")
                                {
                                    sKey = "@B6TreatFrTo@";
                                    sValue = Utils.DateConv(dr["TreatFrDt"] + "", ".") + "~" + Utils.DateConv(dr["TreatToDt"] + "", ".");
                                }
                                rUtil.ReplaceTableRow(oTbl과거력.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
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
                                if (col.ColumnName == "OthClaimReqAmt") sValue = Utils.AddComma(sValue) + "원";
                                if (col.ColumnName == "OthPayAmt") sValue = Utils.AddComma(sValue) + "원";
                                rUtil.ReplaceTableRow(oTbl타사가입.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        TableRow oRow = null;
                        string cd = "";

                        //①약관 청약서 전달 여부
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4TermsRelayCnts@");
                        cd = Utils.ConvertToString(dr["TermsRelayCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "X") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4TermsRelayCnts@", dr["TermsRelayCnts"] + "");

                        //②청약녹취 이행,자필서명 여부
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4HandSignCnts@");
                        cd = Utils.ConvertToString(dr["HandSignCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "X") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4HandSignCnts@", dr["HandSignCnts"] + "");

                        //③상품판매시 면책약관 설명 여부
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ExTermsExplnCnts@");
                        cd = Utils.ConvertToString(dr["ExTermsExplnCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "3") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ExTermsExplnCnts@", dr["ExTermsExplnCnts"] + "");

                        //④면책약관 적용의 합리성
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ExTermsAplyCnts@");
                        cd = Utils.ConvertToString(dr["ExTermsAplyCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "3") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ExTermsAplyCnts@", dr["ExTermsAplyCnts"] + "");

                        //⑤면/부책 판단 구비서류 적정성
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ReqDocuAdeqCnts@");
                        cd = Utils.ConvertToString(dr["ReqDocuAdeqCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "3") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ReqDocuAdeqCnts@", dr["ReqDocuAdeqCnts"] + "");

                        //⑥면책약관의 이해도
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ExTermsUndstdLvlCnts@");
                        cd = Utils.ConvertToString(dr["ExTermsUndstdLvlCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(2), "□", "■");
                        if (cd == "3") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ExTermsUndstdLvlCnts@", dr["ExTermsUndstdLvlCnts"] + "");

                        //⑦작성자 불이익의 원칙 적용
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4WriterDisadvantageCnts@");
                        cd = Utils.ConvertToString(dr["WriterDisadvantageYn"]);
                        if (cd == "Y") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "N") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4WriterDisadvantageCnts@", dr["WriterDisadvantageCnts"] + "");

                        //⑧관련 판례 및 조정례 검토
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4CaseLawAdjReviewCnts@");
                        cd = Utils.ConvertToString(dr["CaseLawAdjReviewYn"]);
                        if (cd == "Y") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "N") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4CaseLawAdjReviewCnts@", dr["CaseLawAdjReviewCnts"] + "");

                        //⑨조검토 가능성 요소
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ReviewPossibleCnts@");
                        cd = Utils.ConvertToString(dr["ReviewPossibleYn"]);
                        if (cd == "Y") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "N") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ReviewPossibleCnts@", dr["ReviewPossibleCnts"] + "");

                        //⑩종합의견
                        oRow = rUtil.GetTableRow(oTbl체크리스트?.Elements<TableRow>(), "@B4ExTermsTtlOpinionCnts@");
                        cd = Utils.ConvertToString(dr["ExTermsTtlOpinionCode"]);
                        if (cd == "1") rUtil.SetText(oRow.GetCell(1), "□", "■");
                        if (cd == "2") rUtil.SetText(oRow.GetCell(3), "□", "■");
                        rUtil.ReplaceTableRow(oRow, "@B4ExTermsTtlOpinionCnts@", dr["ExTermsTtlOpinionCnts"] + "");
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
                                if (col.ColumnName == "GuideDt") sValue = Utils.DateConv(sValue, ".");
                                rUtil.ReplaceTableRow(oTbl민원예방.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
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

                    dtB = pds.Tables["DataBlock11"];
                    sPrefix = "B11";
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

                    dtB = pds.Tables["DataBlock12"];
                    sPrefix = "B12";
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
