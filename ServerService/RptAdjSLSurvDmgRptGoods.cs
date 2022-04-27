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
    public class RptAdjSLSurvDmgRptGoods
    {
        private string myPath = Application.StartupPath;

        public RptAdjSLSurvDmgRptGoods(string path)
        {
            this.myPath = path;
        }

        public Response GetReport(ReportParam para, ref string rptPath, ref string rptName)
        {
            try
            {
                YLWService.YlwSecurityJson security = YLWService.YLWServiceModule.SecurityJson.Clone();  //깊은복사
                security.serviceId = "Metro.Package.AdjSL.BisRprtGoodsPrint";
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

                string sSampleXSD = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물).xsd";
                string sSampleAddFile = "";
                List<string> addFiles = new List<string>();

                DataSet pds = new DataSet();
                pds.ReadXml(sSampleXSD);
                string xml = yds.GetXml();
                using (XmlReader xmlReader = XmlReader.Create(new StringReader(xml)))
                {
                    pds.ReadXml(xmlReader);
                }

                string sSampleDocx = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Head.docx";
                string sSample1Relt = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvDmgRptGoods_Head toHead = new RptAdjSLSurvDmgRptGoods_Head();
                string sRet = toHead.SetSample1(sSampleDocx, sSampleXSD, pds, sSample1Relt);
                if (sRet != "")
                {
                    return new Response() { Result = -1, Message = sRet };
                }

                string[] arryStartWord = new string[] { "가", "나", "다", "라", "마", "바", "사", "아", "자", "차", "카", "타", "파", "하", "가가", "가나", "가다", "가라", "가마", "가바", "가사", "가아", "가자", "가차", "가카", "가타", "가파", "가하" };

                DataTable dtB = pds.Tables["DataBlock3"];
                dtB.DefaultView.Sort = "ObjCatgCd ASC";
                dtB = dtB.DefaultView.ToTable();
                if (dtB != null)
                {
                    for (int ii = 0; ii < dtB.Rows.Count; ii++)
                    {
                        int ObjCatgCd = Utils.ToInt(dtB.Rows[ii]["ObjCatgCd"]);
                        int ObjSeq = Utils.ToInt(dtB.Rows[ii]["ObjSeq"]);

                        sSampleXSD = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Object.xsd";
                        DataSet tds = new DataSet();
                        tds.ReadXml(sSampleXSD);
                        tds.Tables["DataBlock3"].ImportRow(dtB.Rows[ii]);
                        foreach (DataRow drTmp in pds.Tables["DataBlock4"].Select("ObjCatgCd = " + ObjCatgCd + " AND ObjSeq = " + ObjSeq + " "))
                        {
                            tds.Tables["DataBlock4"].ImportRow(drTmp);
                        }
                        foreach (DataRow drTmp in pds.Tables["DataBlock5"].Select("ObjCatgCd = " + ObjCatgCd + " AND ObjSeq = " + ObjSeq + " "))
                        {
                            tds.Tables["DataBlock5"].ImportRow(drTmp);
                        }

                        if (!tds.Tables["DataBlock3"].Columns.Contains("ObjCatgCdNm")) tds.Tables["DataBlock3"].Columns.Add("ObjCatgCdNm");
                        DataRow drT = tds.Tables["DataBlock3"].Rows[0];
                        drT["ObjCatgCdNm"] = arryStartWord[ii] + ". " + drT["InsurObjDvs"];
                        if (ObjCatgCd % 10 == 1 || ObjCatgCd % 10 == 2)  //건물,시설물
                        {
                            if (ObjCatgCd % 10 == 1) drT["ObjCatgCdNm"] += " - 건물";
                            if (ObjCatgCd % 10 == 2) drT["ObjCatgCdNm"] += " - 시설물";
                            sSampleDocx = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Building.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            RptAdjSLSurvDmgRptGoods_Building toWord = new RptAdjSLSurvDmgRptGoods_Building();
                            sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, tds, sSampleAddFile);
                            if (sRet != "")
                            {
                                return new Response() { Result = -1, Message = sRet };
                            }
                            addFiles.Add(sSampleAddFile);
                        }
                        else if (ObjCatgCd % 10 == 3 || ObjCatgCd % 10 == 4)  //기계기구,집기비품
                        {
                            if (ObjCatgCd % 10 == 3) drT["ObjCatgCdNm"] += " - 기계/기구";
                            if (ObjCatgCd % 10 == 4) drT["ObjCatgCdNm"] += " - 집기/비품";
                            sSampleDocx = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Object.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            RptAdjSLSurvDmgRptGoods_Object toWord = new RptAdjSLSurvDmgRptGoods_Object();
                            sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, tds, sSampleAddFile);
                            if (sRet != "")
                            {
                                return new Response() { Result = -1, Message = sRet };
                            }
                            addFiles.Add(sSampleAddFile);
                        }
                        else if (ObjCatgCd % 10 == 5 || ObjCatgCd % 10 == 6)  //재고자산,기타
                        {
                            if (ObjCatgCd % 10 == 5) drT["ObjCatgCdNm"] += " - 재고자산";
                            if (ObjCatgCd % 10 == 6) drT["ObjCatgCdNm"] += " - 기타품목";
                            sSampleDocx = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Inventory.docx";
                            sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                            RptAdjSLSurvDmgRptGoods_Inventory toWord = new RptAdjSLSurvDmgRptGoods_Inventory();
                            sRet = toWord.SetSample1(sSampleDocx, sSampleXSD, tds, sSampleAddFile);
                            if (sRet != "")
                            {
                                return new Response() { Result = -1, Message = sRet };
                            }
                            addFiles.Add(sSampleAddFile);
                        }
                    }
                }
                sSampleDocx = myPath + @"\보고서\출력설계_2541_서식_손해사정서(재물)_Tail.docx";
                sSampleAddFile = myPath + @"\보고서\Temp\" + Guid.NewGuid().ToString() + ".docx";
                RptAdjSLSurvDmgRptGoods_Tail toTail = new RptAdjSLSurvDmgRptGoods_Tail();
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
                dtB = pds.Tables["DataBlock2"];
                if (dtB != null && dtB.Rows.Count > 0)
                {
                    sfilename = Utils.ConvertToString(dtB.Rows[0]["InsurPrdt"]) + "_" + Utils.ConvertToString(dtB.Rows[0]["Insured"]);
                }
                rptName = "손해사정서_재물(" + sfilename + ").docx";
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
