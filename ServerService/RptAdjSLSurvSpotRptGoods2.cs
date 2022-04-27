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
    public class RptAdjSLSurvSpotRptGoods2
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvSpotRptGoods2(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtGoodsPrintSpot";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2512_서식_현장보고서(재물-대물).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_2512_서식_현장보고서(재물-대물).docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                string sRet = SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                DataTable dtBT = pds.Tables["DataBlock10"];
                if (dtBT != null && dtBT.Rows.Count > 0)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2512_서식_현장보고서(재물-대물)_Pict.docx";
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
                DataTable dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "현장보고서_재물-대물(" + sfilename + ").docx";
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
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B3OthInsurCo@");

                    dtB = pds.Tables["DataBlock3"];
                    if (dtB != null)
                    {
                        //2.보험계약사항 - 타보험 계약사항
                        if (oTbl타보험계약사항 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl타보험계약사항, 2, 2, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "FileNo");
                        Table oTableB = rUtil.GetTable(lstTable, sKey);
                        if (oTableB != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableB, 0, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock6"];
                    sPrefix = "B6";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "RprtCnts");
                        Table oTableC = rUtil.GetTable(lstTable, sKey);
                        if (oTableC != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTableC, 1, dtB.Rows.Count - 1);
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
                    //sKey = rUtil.GetFieldName("B1", "Insured"); 
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B3OthInsurCo@");
                    Table oTableA = rUtil.GetTable(lstTable, "@db99ObjInsurRegsAmtTotal@"); // 3.손해내용
                    TableRow oTblA_1 = rUtil.GetTableRow(oTableA?.Elements<TableRow>(), "@db99ObjInsurRegsAmtTotal@"); //3.손해내용-보상한도액
                    TableRow oTblA_2 = rUtil.GetTableRow(oTableA?.Elements<TableRow>(), "@db99EvatStdLosCntsTotal@"); //3.손해내용-추정손해액
                    TableRow oTblA_3 = rUtil.GetTableRow(oTableA?.Elements<TableRow>(), "@db99ObjSelfBearAmtTotal@"); //3.손해내용-자기부담금
                    TableRow oTblA_4 = rUtil.GetTableRow(oTableA?.Elements<TableRow>(), "@db99ObjGivInsurAmtTotal@"); //3.손해내용-추정지급보험금
                    
                    Table oTableB = rUtil.GetTable(lstTable, "@B5FileNo@");
                    Table oTableC = rUtil.GetTable(lstTable, "@B6RprtCnts@");

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
                    }

                    double db2ObjInsurRegsAmt = 0;
                    double db2ObjInsValueTot = 0;
                    double db2ObjTotAmt = 0;
                    double db2ObjGivInsurAmt = 0;
                    string db2EvatStdLosCnts = "";
                    double db2ObjRmnAmt = 0;
                    /*
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
                                if (col.ColumnName == "ObjInsurRegsAmt")
                                {
                                    db2ObjInsurRegsAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjInsValueTot")
                                {
                                    db2ObjInsValueTot += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjTotAmt")
                                {
                                    db2ObjTotAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjGivInsurAmt")
                                {
                                    db2ObjGivInsurAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "EvatStdLosCnts")
                                {
                                    if (db2EvatStdLosCnts != "") db2EvatStdLosCnts += "\n";
                                    db2EvatStdLosCnts += sValue;
                                }
                                if (col.ColumnName == "ObjRmnAmt")
                                {
                                    db2ObjRmnAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTableA.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }*/
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db2ObjInsurRegsAmt@", Utils.AddComma(db2ObjInsurRegsAmt));
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db2ObjInsValueTot@", Utils.AddComma(db2ObjInsValueTot));
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db2ObjTotAmt@", Utils.AddComma(db2ObjTotAmt));
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db2ObjGivInsurAmt@", Utils.AddComma(db2ObjGivInsurAmt));
                    rUtil.ReplaceTables(lstTable, "@db2EvatStdLosCnts@", db2EvatStdLosCnts);
                    rUtil.ReplaceTables(lstTable, "@db2ObjRmnAmt@", Utils.AddComma(db2ObjRmnAmt));

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (oTbl타보험계약사항 != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = (i + 1) * 2;
                                foreach (DataColumn col in dtB.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "OthCtrtDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "OthCtrtExprDt") sValue = Utils.DateConv(sValue, ".");
                                    if (col.ColumnName == "OthInsurRegsAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "OthSelfBearAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 1), sKey, sValue);
                                }
                            }
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
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
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
                                rUtil.ReplaceTableRow(oTableB.GetRow(i), sKey, sValue);
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
                                rUtil.ReplaceTableRow(oTableC.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }

                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjInsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "EvatStdLosCnts") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjGivInsurAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];
                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "ObjInsurRegsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "EvatStdLosCnts") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "ObjGivInsurAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    //합계
                    double db99ObjInsurRegsAmtTotal = 0; //보상한도액
                    double db99EvatStdLosCntsTotal = 0; //추정손해액
                    double db99ObjSelfBearAmtTotal = 0; //자기부담금
                    double db99ObjGivInsurAmtTotal = 0; //추정지급보험금

                    DataTable dtB_7 = pds.Tables["DataBlock7"];
                    DataTable dtB_8 = pds.Tables["DataBlock8"];
                    if (dtB_7.Rows.Count < 1) dtB_7.Rows.Add();
                    if (dtB_8.Rows.Count < 1) dtB_8.Rows.Add();
                    DataRow dr_7 = dtB_7.Rows[0];
                    DataRow dr_8 = dtB_8.Rows[0];

                    //보상한도액-합계
                    db99ObjInsurRegsAmtTotal = Utils.ToDouble(dr_7["ObjInsurRegsAmt"]) + Utils.ToDouble(dr_8["ObjInsurRegsAmt"]);
                    //추정손해액-합계
                    db99EvatStdLosCntsTotal = Utils.ToDouble(dr_7["EvatStdLosCnts"]) + Utils.ToDouble(dr_8["EvatStdLosCnts"]);
                    //자기부담금-합계
                    db99ObjSelfBearAmtTotal = Utils.ToDouble(dr_7["ObjSelfBearAmt"]) + Utils.ToDouble(dr_8["ObjSelfBearAmt"]);
                    //추정지급보험금-합계
                    db99ObjGivInsurAmtTotal = Utils.ToDouble(dr_7["ObjGivInsurAmt"]) + Utils.ToDouble(dr_8["ObjGivInsurAmt"]);

                    rUtil.ReplaceTableRow(oTblA_1, "@db99ObjInsurRegsAmtTotal@", Utils.AddComma(db99ObjInsurRegsAmtTotal)); //보상한도액-합계
                    rUtil.ReplaceTableRow(oTblA_2, "@db99EvatStdLosCntsTotal@", Utils.AddComma(db99EvatStdLosCntsTotal)); //추정손해액-합계
                    rUtil.ReplaceTableRow(oTblA_3, "@db99ObjSelfBearAmtTotal@", Utils.AddComma(db99ObjSelfBearAmtTotal)); //자기부담금-합계
                    rUtil.ReplaceTableRow(oTblA_4, "@db99ObjGivInsurAmtTotal@", Utils.AddComma(db99ObjGivInsurAmtTotal)); //추정지급보험금-합계

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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");

                    dtB = pds.Tables["DataBlock10"];
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
                    Table oTbl현장사진 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");

                    dtB = pds.Tables["DataBlock10"];
                    sPrefix = "B10";
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
