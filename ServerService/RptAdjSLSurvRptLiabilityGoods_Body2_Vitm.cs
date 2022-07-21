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
    public class RptAdjSLSurvRptLiabilityGoods_Body2_Vitm
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile, DataRow pdr7, DataTable pdt10, string num)
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");

                    sPrefix = "B10";
                    if (pdt10 != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((pdt10.Rows.Count + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                rUtil.TableAddRow(oTbl손해상황, 2, 1);
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B10AcdtPictImage@");

                    rUtil.ReplaceTextAllParagraph(doc, "@B7Num@", num);
                    rUtil.ReplaceTextAllParagraph(doc, "@B7Title@", pdr7["Vitm"] + "");

                    dtB = pdt10;
                    sPrefix = "B10";
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
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
                                TableRow xrow1 = oTbl손해상황.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2500000L, 2000000L);
                                }
                                catch { }

                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = dr["AcdtPictCnts"] + "";
                                TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
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
