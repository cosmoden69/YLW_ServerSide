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
    public class RptAdjSLSurvRptLiabilityGoods_Car_Head1
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityGoods_Car_Head1()
        {
        }

        public RptAdjSLSurvRptLiabilityGoods_Car_Head1(string path)
        {
            this.myPath = path;
        }

        internal string SetSample1(string sSampleDocx, string sSampleXSD, DataSet pds, string sSample1Relt)
        {
            throw new NotImplementedException();
        }

        /*
public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
{
   try
   {
       YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
       security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintGoodsCar";
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

       string sSampleXSD = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량).xsd";

       DataSet pds = new DataSet();
       pds.ReadXml(sSampleXSD);
       string xml = yds.GetXml();
       using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
       {
           pds.ReadXml(xmlReader);
       }

       string sSample1Docx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량).docx";
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
       rptName = "종결보고서_배책_대물_차량(" + sfilename + ").docx";
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
*/
        private string SetSample11(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    Table oTbl피해물관련사항 = rUtil.GetTable(lstTable, "@B5DmobSortNo@");
                    //Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B7OthInsurCo@");
                    //Table oTbl피보험자사항_법인 = rUtil.GetTable(lstTable, "@B1IsrdAgrmAgt@");
                    //Table oTbl피보험자사항_개인 = rUtil.GetTable(lstTable, "@B1IsrdJob@");

                    /*
                    dtB = pds.Tables["DataBlock1"];
                    if (dtB != null)
                    {
                        DataRow dr = dtB.Rows[0];
                        if ((sValue != null)&&(Utils.ToInt(dr["InsuredFg"])%10 == 2)) //법인
                        {
                            oTbl피보험자사항_개인.Remove();
                        }
                        else //개인
                        {
                            oTbl피보험자사항_법인.Remove();
                        }
                    }
                    */


                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "DmobSortNo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "InsurGivObj");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }



                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        sKey = rUtil.GetFieldName(sPrefix, "OthInsurCo");
                        Table oTable = rUtil.GetTable(lstTable, sKey);
                        if (oTable != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTable, 2, 2, dtB.Rows.Count - 1);
                            
                        }
                    }



                    dtB = pds.Tables["DataBlock5"];
                    sPrefix = "B5";
                    if (dtB != null)
                    {
                        if (oTbl피해물관련사항 != null)
                        {
                            //테이블의 끝에 추가
                            //rUtil.TableAddRow(oTbl피해물관련사항, 1, dtB.Rows.Count - 1);
                            rUtil.TableInsertRows(oTbl피해물관련사항, 0, 9, 1);
                        }
                    }


                    /*
                         dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
                    if (dtB != null)
                    {
                        if (oTableH != null)
                        {
                            double cnt = dtB.Rows.Count;
                            for (int i = 1; i < cnt; i++)
                            {
                                //테이블의 끝에 추가
                                rUtil.TableInsertRows(oTableH, 1, 7, 1);
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
                     */






                    /*
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        sKey = "@B1AcdtCnts@";
                        Table oTblA = rUtil.GetTable(lstTable, sKey);
                        sKey = "@B2AcdtPictImage1@";
                        TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                        Table oTableA = oTblARow?.GetCell(1).Elements<Table>().FirstOrDefault();
                        if (oTableA != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = dtB.Rows.Count;//Math.Truncate((dtB.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTableA, 0, 1);
                                rUtil.TableAddRow(oTableA, 1, 1);
                            }
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
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTable, 1, dtB.Rows.Count - 1);
                        }
                    }

                    
                    */
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
                    Table oTable총괄표 = rUtil.GetTable(lstTable, "@B2DmobSortNo@");
                    Table oTable보험금지급처 = rUtil.GetTable(lstTable, "@B3InsurGivObj@");
                    Table oTbl타보험계약사항 = rUtil.GetTable(lstTable, "@B4OthInsurCo@");
                    //Table oTableC = rUtil.GetTable(lstTable, "@B8ExpsReqAmt1@");


                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    sKey = "@B1AcdtCnts@";
                    Table oTblA = rUtil.GetTable(lstTable, sKey);
                    /*
                    sKey = "@B2AcdtPictImage1@";
                    TableRow oTblARow = rUtil.GetTableRow(oTblA?.Elements<TableRow>(), sKey);
                    Table oTableA = oTblARow?.GetCell(1).Elements<Table>().FirstOrDefault();
                    */


                    //sKey = "@B1AcdtCnts@";
                    //Table oTblB = rUtil.GetTable(lstTable, sKey);
                    //sKey = "@B2AcdtPictImage2@";
                    //TableRow oTblBRow = rUtil.GetTableRow(oTblB?.Elements<TableRow>(), sKey);
                    //Table oTableB = oTblBRow?.GetCell(1).Elements<Table>().FirstOrDefault();
/*
                    dtB = pds.Tables["DataBlock1"];
                    sPrefix = "B1";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (!dtB.Columns.Contains("DoOthExpsHedText")) dtB.Columns.Add("DoOthExpsHedText");
                        {
                            if (Utils.ConvertToString(dr["DoOthExpsHed"]) == "")
                            {
                                dr["DoOthExpsHedText"] = "4. ";
                            }
                            else
                            {
                                dr["DoOthExpsHedText"] = "4." + dr["DoOthExpsHed"];
                            }
                        }

                        if (!dtB.Columns.Contains("DoOthExpsHedText")) dtB.Columns.Add("DoOthExpsHedText");
                        {
                            if ((Utils.ConvertToInt(dr["DoOthExpsReq"]) == 0) && (Utils.ConvertToString(dr["DoOthExpsReq"]) == "") && (Utils.ConvertToInt(dr["DoOthExpsAmt"]) == 0) && (Utils.ConvertToString(dr["DoOthExpsAmt"]) == ""))
                            {
                                dr["DoOthExpsHedText"] = " ";
                                dr["DoOthExpsReq"] = 0;
                                dr["DoOthExpsAmt"] = 0;
                                dr["DoOthExpsCmnt"] = " ";
                                dr["DoOthExpsBss"] = " ";
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
                            if (col.ColumnName == "FldRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "MidRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "LasRptSbmsDt") sValue = Utils.DateFormat(sValue, "yyyy년 MM월 dd일");
                            if (col.ColumnName == "EmpPhone") sValue = (sValue == "" ? "" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "IsrtTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "IsrdTel") sValue = (sValue == "" ? "-" : Utils.TelNumber(sValue));
                            if (col.ColumnName == "CtrtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd") + " ~";
                            if (col.ColumnName == "CtrtExprDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "FixFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FixToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            //if (col.ColumnName == "InsurValue") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "InsurRegsAmt2") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotReq") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue); 수정
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AgrmAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoBivInsurAmt") sValue = Utils.AddComma(sValue);
                            //if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue); 삭제
                            if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "InsurRegsAmtRevw") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "SelfBearAmtRevw") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AcdtDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "AcdtTm") sValue = Utils.TimeFormat(sValue, "HH:mm");
                            if (col.ColumnName == "LeadAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "ChrgAdjuster") sValue = Utils.Adjuster(sValue);
                            if (col.ColumnName == "DoFixReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoFixAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNoCarfeeAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoRentCarAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoOthExpsAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSubTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSubTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoNglgBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoTotAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearReq") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoSelfBearAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "DoGivInsurReq") sValue = Utils.AddComma(sValue);
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

                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
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
                                if (col.ColumnName == "CarTyp") if (sValue != "") { sValue = "(" + sValue + ")"; } //차량종류
                                if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue) + "원"; //보상한도액
                                if (col.ColumnName == "ReqAmt") sValue = Utils.AddComma(sValue) + "원"; //청구액
                                if (col.ColumnName == "DoLosAmt") sValue = Utils.AddComma(sValue) + "원"; //손해액
                                if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue) + "원"; //자기부담금
                                if (col.ColumnName == "GivInsurAmt") sValue = Utils.AddComma(sValue) + "원"; //지급보험금  
                                rUtil.ReplaceTableRow(oTable총괄표.GetRow(i + 1), sKey, sValue);
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
                                if (col.ColumnName == "GivObjRegno") if (sValue != "") { sValue = "(" + sValue + ")"; } //주민번호(사업자번호)
                                if (col.ColumnName == "GivObjInsurAmt") sValue = Utils.AddComma(sValue) + "원   "; //지급보험금
                                rUtil.ReplaceTableRow(oTable보험금지급처.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }



                    dtB = pds.Tables["DataBlock4"];
                    sPrefix = "B4";
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
                                    if (col.ColumnName == "OthInsurRegsAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; }
                                    if (col.ColumnName == "OthSelfBearAmt") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "원"; }
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 0), sKey, sValue);
                                    rUtil.ReplaceTableRow(oTbl타보험계약사항.GetRow(rnum + 1), sKey, sValue);
                                }
                            }
                        }
                    }

                    */
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
                                if (col.ColumnName == "VitmNglgRate") if (Utils.ConvertToInt(sValue) > 0) { sValue = Utils.AddComma(sValue) + "%"; }
                                rUtil.ReplaceTables(lstTable, sKey, sValue);
                                //rUtil.ReplaceTableRow(oTable총괄표.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }
                    /*
                    dtB = pds.Tables["DataBlock2"];
                    sPrefix = "B2";
                    if (dtB != null)
                    {
                        if (oTableA != null)
                        {
                            if (dtB.Rows.Count < 1) dtB.Rows.Add();
                            for (int i = 0; i < dtB.Rows.Count; i++)
                            {
                                DataRow dr = dtB.Rows[i];
                                int rnum = i * 2;
                                int cnum = 0;
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage1");
                                sValue = dr["AcdtPictImage1"] + "";
                                TableRow xrow1 = oTableA.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(cnum), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(cnum), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts1");
                                sValue = dr["AcdtPictCnts1"] + "";
                                TableRow xrow2 = oTableA.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(cnum), sKey, sValue);

                                //--------------------------------------------------------------------------------------------------

                                cnum = 1;
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage2");
                                sValue = dr["AcdtPictImage2"] + "";
                                rUtil.SetText(xrow1.GetCell(cnum), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(cnum), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts2");
                                sValue = dr["AcdtPictCnts2"] + "";
                                rUtil.SetText(xrow2.GetCell(cnum), sKey, sValue);
                            }
                        }
                    }
                    
                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        if (!dtB.Columns.Contains("VitmNglgRatePer")) dtB.Columns.Add("VitmNglgRatePer");
                        dr["VitmNglgRatePer"] = dr["VitmNglgRate"] + "%";


                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";

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
                            if (col.ColumnName == "FixFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            if (col.ColumnName == "FixToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
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
                    

                    //손해액 산정내역
                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    if (dtB != null)
                    {
                        //1.수리비
                        DataRow[] drs = dtB?.Select("ExpsGrp = 1");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        double dReq = 0;
                        double dAmt = 0;
                        string sEvatRslt = "";
                        string sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";

                        }
                        TableRow oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt1@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt1@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt1@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt1@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss1@", sExpsBss);
                        }

                        //2.휴차료
                        drs = dtB?.Select("ExpsGrp = 2");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt2@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt2@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt2@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt2@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss2@", sExpsBss);
                        }

                        //3.대차료
                        drs = dtB?.Select("ExpsGrp = 3");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt3@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt3@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt3@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt3@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss3@", sExpsBss);
                        }

                        //4.기타비용
                        drs = dtB?.Select("ExpsGrp = 4");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt4@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt4@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt4@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt4@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss4@", sExpsBss);
                        }

                        //*소계
                        drs = dtB?.Select("ExpsGrp = 91");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt91@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt91@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt91@", Utils.AddComma(dAmt));
                        }

                        //5.과실부담금
                        drs = dtB?.Select("ExpsGrp = 5");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt5@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt5@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt5@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt5@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss5@", sExpsBss);
                        }

                        //*합계
                        drs = dtB?.Select("ExpsGrp = 92");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt92@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt92@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt92@", Utils.AddComma(dAmt));
                        }

                        //6.자기부담금
                        drs = dtB?.Select("ExpsGrp = 6");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt6@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt6@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt6@", Utils.AddComma(dAmt));
                            rUtil.ReplaceTableRow(oRow, "@B8EvatRslt6@", sEvatRslt);
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsBss6@", sExpsBss);
                        }

                        //*예상지급보험금
                        drs = dtB?.Select("ExpsGrp = 93");
                        if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                        dReq = 0;
                        dAmt = 0;
                        sEvatRslt = "";
                        sExpsBss = "";
                        for (int i = 0; i < drs.Length; i++)
                        {
                            dReq += Utils.ToDouble(drs[i]["ExpsReqAmt"] + "");
                            dAmt += Utils.ToDouble(drs[i]["ExpsDoLosAmt"] + "");
                            sEvatRslt = drs[i]["EvatRslt"] + "";
                            sExpsBss = drs[i]["ExpsBss"] + "";
                        }
                        oRow = rUtil.GetTableRow(oTableC?.Elements<TableRow>(), "@B8ExpsReqAmt93@");
                        if (oRow != null)
                        {
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsReqAmt93@", Utils.AddComma(dReq));
                            rUtil.ReplaceTableRow(oRow, "@B8ExpsDoLosAmt93@", Utils.AddComma(dAmt));
                        }
                    }
                    */


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
