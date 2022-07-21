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
    public class RptAdjSLSurvMidRptLiabilityGoods
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvMidRptLiabilityGoods(string path)
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물).xsd";
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
                //string sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Head.docx";
                string sSampleDocx = myPath + @"\보고서\출력설계_2534_서식_중간보고서(배책-대물)_Head.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Head toHead = new RptAdjSLSurvRptLiabilityGoods_Head();
                string sRet = toHead.SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string[] arryStartWord = new string[] { "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", "가가", "가나", "가다", "가라", "가마", "가바", "가사", "가아", "가자", "가차", "가카", "가타", "가파", "가하" };
                string[] arryNumbrWord = new string[] { "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮" };
                DataTable dtB = null;
                DataRow[] drs = null;

                //계약 및 사고관련사항 - 증권
                dtB = pds.Tables["DataBlock4"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Head_Insur.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Head_Insur toHead1 = new RptAdjSLSurvRptLiabilityGoods_Head_Insur();
                    sRet = toHead1.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dtB.Rows[i], "증권 " + arryNumbrWord[i]);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(false);
                }

                //계약 및 사고관련사항 - 타보험 계약사항
                sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Head_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Head_Tail toBody = new RptAdjSLSurvRptLiabilityGoods_Head_Tail();
                sRet = toBody.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(false);

                //일반사항 - 피해자 및 피해현황
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Head_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Head_Vitm toHead2 = new RptAdjSLSurvRptLiabilityGoods_Head_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock8"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt8 = (drs.Length < 1 ? pds.Tables["DataBlock8"].Clone() : drs.CopyToDataTable());
                    sRet = toHead2.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt8);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                //사고사항 - 사고개요
                sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body1.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Body1 toBody1 = new RptAdjSLSurvRptLiabilityGoods_Body1();
                sRet = toBody1.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
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
                    sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body1_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Body1_Vitm toBody1V = new RptAdjSLSurvRptLiabilityGoods_Body1_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock9"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt9 = (drs.Length < 1 ? pds.Tables["DataBlock9"].Clone() : drs.CopyToDataTable());
                    sRet = toBody1V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt9);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                //사고사항 - 피해사진
                sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body2.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Body2 toBody2 = new RptAdjSLSurvRptLiabilityGoods_Body2();
                sRet = toBody2.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //사고사항 - 피해 사진
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body2_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Body2_Vitm toBody2V = new RptAdjSLSurvRptLiabilityGoods_Body2_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock10"]?.Select("VitmSubSeq = " + vitmSubSeq + " ", "AcdtPictSerl");
                    DataTable dt10 = (drs.Length < 1 ? pds.Tables["DataBlock10"].Clone() : drs.CopyToDataTable());
                    sRet = toBody2V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt10, Utils.ConvertToString(i + 1));
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);
                }

                //법률상 배상책임 성립 여부
                sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body3.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Body3 toBody3 = new RptAdjSLSurvRptLiabilityGoods_Body3();
                sRet = toBody3.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }
                addFiles.Add(sSampleAddFile);
                addNewLine.Add(true);

                //손해사정 - 피해자
                dtB = pds.Tables["DataBlock7"];
                if (dtB.Rows.Count < 1) dtB.Rows.Add();
                for (int i = 0; i < dtB.Rows.Count; i++)
                {
                    sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body3_Vitm.docx";
                    sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                    RptAdjSLSurvRptLiabilityGoods_Body3_Vitm toBody3V = new RptAdjSLSurvRptLiabilityGoods_Body3_Vitm();

                    DataRow dr7 = dtB.Rows[i];
                    string vitmSubSeq = Utils.ConvertToString(dr7["VitmSubSeq"]);
                    drs = pds.Tables["DataBlock11"]?.Select("VitmSubSeq = " + vitmSubSeq + " ");
                    DataTable dt11 = (drs.Length < 1 ? pds.Tables["DataBlock11"].Clone() : drs.CopyToDataTable());
                    sRet = toBody3V.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, dr7, dt11);
                    if (sRet != "")
                    {
                        return new Response() { Result = -1, Message = sRet };
                    }
                    addFiles.Add(sSampleAddFile);
                    addNewLine.Add(i == 0 ? false : true);

                    //목적물별 상세
                    for (int j = 0; j < drs.Length; j++)
                    {
                        string dmobDvs = Utils.ConvertToString(drs[j]["DmobDvs"]);
                        if (dmobDvs == "300102010" || dmobDvs == "300102011" || dmobDvs == "300102013")  //집기/비품, 가재도구, 재고자산
                        {
                            sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body3_Vitm_Dmob2.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            RptAdjSLSurvRptLiabilityGoods_Body3_Vitm_Dmob2 toBodyA = new RptAdjSLSurvRptLiabilityGoods_Body3_Vitm_Dmob2();
                            sRet = toBodyA.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, drs[j], arryStartWord[j]);
                        }
                        else
                        {
                            sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Body3_Vitm_Dmob1.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            RptAdjSLSurvRptLiabilityGoods_Body3_Vitm_Dmob1 toBodyA = new RptAdjSLSurvRptLiabilityGoods_Body3_Vitm_Dmob1();
                            sRet = toBodyA.SetSample1(sSampleDocx, sSampleXSD, pds, sSampleAddFile, drs[j], arryStartWord[j]);
                        }

                        if (sRet != "")
                        {
                            return new Response() { Result = -1, Message = sRet };
                        }
                        addFiles.Add(sSampleAddFile);
                        addNewLine.Add(false);
                    }
                }

                //sSampleDocx = myPath + @"\보고서\출력설계_2554_서식_종결보고서(배책-대물)_Tail.docx";
                sSampleDocx = myPath + @"\보고서\출력설계_2534_서식_중간보고서(배책-대물)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvRptLiabilityGoods_Tail toTail = new RptAdjSLSurvRptLiabilityGoods_Tail();
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
                rptName = "중간보고서_배책_대물(" + sfilename + ").docx";
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
