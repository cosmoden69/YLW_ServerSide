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
    public class RptAdjSLSurvRptLiabilityDB
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityDB(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintKB";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                //Head.doc
                string sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Head.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityDB_Head toHead = new RptAdjSLSurvRptLiabilityDB_Head();
                string sRet = toHead.SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string[] arryStartWord = new string[] { "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", "가가", "가나", "가다", "가라", "가마", "가바", "가사", "가아", "가자", "가차", "가카", "가타", "가파", "가하" };
                /*
                var db11Cnt = 0;
                DataTable dtB = pds.Tables["DataBlock11"];
                db11Cnt = dtB.Rows.Count;
                */
                DataTable dtB = pds.Tables["DataBlock1"];
                //DataTable dtB = pds.Tables["DataBlock1"];
                dr = dtB.Rows[0];


                //AcptMgmt(대인만 존재:572, 대물만 존재:575, 대인+대물 동시에 존재:573)

                //대인만 존재할 경우(Pers.doc)
                if ((dr["DIGivInsurAmt"] != null || Utils.ConvertToString(dr["DIGivInsurAmt"]) != "")  // 대인:산정손해액-예상지급보험금은 있고
                    && (dr["DoGivInsurAmt"] == null || Utils.ConvertToString(dr["DoGivInsurAmt"]) == "")) // && 대물:산정손해액 - 예상지급보험금은 없는경우
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Pers.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityDB_Pers toWord = new RptAdjSLSurvRptLiabilityDB_Pers();
                    sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);

                }
                //대물만 존재할 경우(Goods.doc)
                else if ((dr["DoGivInsurAmt"] != null || Utils.ConvertToString(dr["DoGivInsurAmt"]) != "")  // 대물:산정손해액-예상지급보험금은 있고
                && (dr["DIGivInsurAmt"] == null || Utils.ConvertToString(dr["DIGivInsurAmt"]) == "")) // && 대인:산정손해액 - 예상지급보험금은 없는경우
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Goods.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityDB_Goods toWord = new RptAdjSLSurvRptLiabilityDB_Goods();
                    sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }

                //대인과 대물 둘 다 존재할 경우
                else if ((dr["DoGivInsurAmt"] != null || Utils.ConvertToString(dr["DoGivInsurAmt"]) != "")  // 대물:산정손해액-예상지급보험금도 있고
                    && (dr["DIGivInsurAmt"] != null || Utils.ConvertToString(dr["DIGivInsurAmt"]) != "")) // && 대인:산정손해액 - 예상지급보험금도 있는경우
                {
                    //Pers.doc
                    {
                        sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Pers.docx";
                        sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                        RptAdjSLSurvRptLiabilityDB_Pers toWord = new RptAdjSLSurvRptLiabilityDB_Pers();
                        sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                        if (sRet != "")
                        {
                            return new Response() { Result = -1, Message = sRet };
                        }
                        addFiles.Add(sSampleAddFile);
                    }
                    
                    //Goods.doc
                    {
                        sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Goods.docx";
                        sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                        RptAdjSLSurvRptLiabilityDB_Goods toWord = new RptAdjSLSurvRptLiabilityDB_Goods();
                        sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                        if (sRet != "")
                        {
                            return new Response() { Result = -1, Message = sRet };
                        }
                        addFiles.Add(sSampleAddFile);
                    }
                    
                }

                //Tail.doc
                sSampleDocx = myPath + @"\보고서\출력설계_2582_서식_DB_종결보고서(배책)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityDB_Tail toTail = new RptAdjSLSurvRptLiabilityDB_Tail();
                sRet = toTail.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
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
                dtB = pds.Tables["DataBlock1"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "종결보고서_배책_DB(" + sfilename + ").docx";
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
    }
}
