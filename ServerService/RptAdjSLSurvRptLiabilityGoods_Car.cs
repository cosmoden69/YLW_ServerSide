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
    public class RptAdjSLSurvRptLiabilityGoods_Car
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvRptLiabilityGoods_Car(string path)
        {
            this.myPath = path;
        }

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
                string sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Head.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Car_Head toHead = new RptAdjSLSurvRptLiabilityGoods_Car_Head();
                string sRet = toHead.SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string[] arryStartWord = new string[] { "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", "가가", "가나", "가다", "가라", "가마", "가바", "가사", "가아", "가자", "가차", "가카", "가타", "가파", "가하" };
                DataTable dtB = pds.Tables["DataBlock1"];
                dr = dtB.Rows[0];


                DataRow[] drs = null;
                var B5Cnt = 0;
                dtB = pds.Tables["DataBlock5"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                B5Cnt = dtB.Rows.Count;

                //Body1 (Ⅲ.일반사항)
                var ObjSeq1 = 0;
                for (int aa = 0; aa < B5Cnt; aa++)
                {

                    drs = dtB?.Select("");
                    ObjSeq1 = Utils.ToInt(drs[aa]["DmobSeq"]);

                    sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Body1.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Car_Body1 toBody1 = new RptAdjSLSurvRptLiabilityGoods_Car_Body1();
                    sRet = toBody1.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, ObjSeq1);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }
                

                //Body2 (Ⅳ.사고사항-사고내용)
                sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Body2.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Car_Body2 toBody2 = new RptAdjSLSurvRptLiabilityGoods_Car_Body2();
                sRet = toBody2.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);

                
                //Body3 (Ⅳ.사고사항-사고사진)
                var ObjSeq3 = 0;
                for (int aa = 0; aa < B5Cnt; aa++)
                {
                    drs = dtB?.Select("");
                    ObjSeq3 = Utils.ToInt(drs[aa]["DmobSeq"]);

                    sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Body3.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Car_Body3 toBody3 = new RptAdjSLSurvRptLiabilityGoods_Car_Body3();
                    sRet = toBody3.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, ObjSeq3);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }


                //Body4 (Ⅴ.법률상 배상책임 성립여부)
                sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Body4.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Car_Body4 toBody4 = new RptAdjSLSurvRptLiabilityGoods_Car_Body4();
                sRet = toBody4.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);


                //Body5 (Ⅵ.손해액 평가)
                var ObjSeq5 = 0;
                for (int aa = 0; aa < B5Cnt; aa++)
                {
                    drs = dtB?.Select("");
                    ObjSeq5 = Utils.ToInt(drs[aa]["DmobSeq"]);

                    sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Body5.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Car_Body5 toBody5 = new RptAdjSLSurvRptLiabilityGoods_Car_Body5();
                    sRet = toBody5.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, ObjSeq5);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                }


                //Tail.doc
                sSampleDocx = myPath + @"\보고서\출력설계_2555_서식_종결보고서(배책-대물-차량)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Car_Tail toTail = new RptAdjSLSurvRptLiabilityGoods_Car_Tail();
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
    }
}
