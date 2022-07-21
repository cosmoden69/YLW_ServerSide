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
    public class RptAdjSLSurvSpotRptLiabilityGoods_Head2_Vitm
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile, DataRow pdr7, DataTable pdt11)
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
                    Table oTbl구분 = rUtil.GetTable(lstTable, "@B11DmobNm@");
                    Table oTbl손해현황 = rUtil.GetTable(lstTable, "@B11DmobDmgStts@");

                    if (pdt11 != null)
                    {
                        if (oTbl구분 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl구분, 1, pdt11.Rows.Count - 1);
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
                    Table oTbl구분 = rUtil.GetTable(lstTable, "@B11DmobNm@");
                    Table oTbl손해현황 = rUtil.GetTable(lstTable, "@B11DmobDmgStts@");

                    rUtil.ReplaceTextAllParagraph(doc, "@B7Title@", pdr7["Vitm"] + "");

                    string db11DmobDmgStts = "";
                    string db11RmnObjRmvCnts = "";
                    dtB = pdt11;
                    sPrefix = "B11";
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
                                if (col.ColumnName == "InsurRegsAmt") sValue = Utils.AddComma(sValue); //보상한도액
                                if (col.ColumnName == "EstmLosAmt") sValue = Utils.AddComma(sValue);   //추정손해액
                                if (col.ColumnName == "SelfBearAmt") sValue = Utils.AddComma(sValue);  //자기부담금
                                if (col.ColumnName == "GivInsurAmt") sValue = Utils.AddComma(sValue);  //추정 지급보험금
                                if (col.ColumnName == "DmobDmgStts")
                                {
                                    if (i > 0) db11DmobDmgStts += "\n";
                                    db11DmobDmgStts += sValue;
                                }
                                if (col.ColumnName == "RmnObjRmvCnts")
                                {
                                    if (i > 0) db11RmnObjRmvCnts += "\n";
                                    db11RmnObjRmvCnts += sValue;
                                }
                                rUtil.ReplaceTableRow(oTbl구분.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl손해현황, "@B11DmobDmgStts@", db11DmobDmgStts);
                    rUtil.ReplaceTable(oTbl손해현황, "@B11RmnObjRmvCnts@", db11RmnObjRmvCnts);

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
