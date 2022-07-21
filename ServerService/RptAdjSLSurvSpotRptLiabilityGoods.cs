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
    public class RptAdjSLSurvSpotRptLiabilityGoods
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvSpotRptLiabilityGoods(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtLiabilityPrintGoods";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();
                List<bool> addNewLine = new List<bool>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                //Head.doc
                string sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvSpotRptLiabilityGoods_Head toHead = new RptAdjSLSurvSpotRptLiabilityGoods_Head();
                string sRet = toHead.SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string[] arryStartWord = new string[] { "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", "가가", "가나", "가다", "가라", "가마", "가바", "가사", "가아", "가자", "가차", "가카", "가타", "가파", "가하" };
                string[] arryNumbrWord = new string[] { "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮" };
                DataTable dtB = null;
                DataRow[] drs = null;

                //일반사항 - 피해자 및 피해현황
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Head_Vitm toHeadV = new RptAdjSLSurvRptLiabilityGoods_Head_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock8"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt8 = (drs.Length < 1 ? pds.Tables["DataBlock8"].Clone() : drs.CopyToDataTable());
                    sRet = toHeadV.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt8);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(false);
                }

                //손해상황
                sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head2.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvSpotRptLiabilityGoods_Head2 toHead2 = new RptAdjSLSurvSpotRptLiabilityGoods_Head2();
                sRet = toHead2.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //손해상황 - 피해자별
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head2_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvSpotRptLiabilityGoods_Head2_Vitm toHead2V = new RptAdjSLSurvSpotRptLiabilityGoods_Head2_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock11"]?.Select("VitmSubSeq = " + vitmSubSeq + " ");
                    DataTable dt11 = (drs.Length < 1 ? pds.Tables["DataBlock11"].Clone() : drs.CopyToDataTable());
                    sRet = toHead2V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt11);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                //사고사진 - 사고 및 피해장소 평면도
                sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head3.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvSpotRptLiabilityGoods_Head3 toHead3 = new RptAdjSLSurvSpotRptLiabilityGoods_Head3();
                sRet = toHead3.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //사고사항 - 사고 및 피해장소 평면도
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head3_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvSpotRptLiabilityGoods_Head3_Vitm toHead3V = new RptAdjSLSurvSpotRptLiabilityGoods_Head3_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock9"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt9 = (drs.Length < 1 ? pds.Tables["DataBlock9"].Clone() : drs.CopyToDataTable());
                    sRet = toHead3V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt9);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                //사고사진 - 피해사진
                sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head4.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvSpotRptLiabilityGoods_Head4 toHead4 = new RptAdjSLSurvSpotRptLiabilityGoods_Head4();
                sRet = toHead4.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //사고사진 - 피해 사진
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Head4_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvSpotRptLiabilityGoods_Head4_Vitm toHead4V = new RptAdjSLSurvSpotRptLiabilityGoods_Head4_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock10"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt10 = (drs.Length < 1 ? pds.Tables["DataBlock10"].Clone() : drs.CopyToDataTable());
                    sRet = toHead4V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt10, Utils.ConvertToString(i + 1));
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                sSampleDocx = myPath + @"\보고서\출력설계_2514_서식_현장보고서(배책-대물)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvSpotRptLiabilityGoods_Tail toTail = new RptAdjSLSurvSpotRptLiabilityGoods_Tail();
                sRet = toTail.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //DOCX 파일합치기 
                WordprocessingDocument wdoc = WordprocessingDocument.Open(sSample1Relt, true);
                MainDocumentPart mainPart = wdoc.MainDocumentPart;
                for (int ii = 0; ii < addFiles.Count; ii++)
                {
                    string addFile = addFiles[ii];
                    RptUtils.AppendFile(mainPart, addFile, addNewLine[ii]);
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
                rptName = "현장보고서_배책_대물(" + sfilename + ").docx";
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
