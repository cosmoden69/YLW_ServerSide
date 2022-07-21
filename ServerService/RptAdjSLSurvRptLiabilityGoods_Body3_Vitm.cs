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
    public class RptAdjSLSurvRptLiabilityGoods_Body3_Vitm
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
                string vitmSubSeq = Utils.ConvertToString(pdr7["VitmSubSeq"]);
                drs = pds.Tables["DataBlock12"]?.Select("VitmSubSeq = " + vitmSubSeq + " ");
                DataTable dt12 = (drs.Length < 1 ? null : drs.CopyToDataTable());

                System.IO.File.Copy(sDocFile, sWriteFile, true);

                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(sWriteFile, true))
                {
                    MainDocumentPart mDoc = wDoc.MainDocumentPart;
                    Document doc = mDoc.Document;
                    RptUtils rUtil = new RptUtils(mDoc);

                    IEnumerable<Table> lstTable = doc.Body.Elements<Table>();
                    Table oTbl평가방법 = rUtil.GetTable(lstTable, "@B11DmobNm@");
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B12ExpsGrpNm@");

                    if (pdt11 != null)
                    {
                        if (oTbl평가방법 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl평가방법, 0, pdt11.Rows.Count - 1);
                        }
                    }

                    if (dt12 != null)
                    {
                        if (oTbl총괄표 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTbl총괄표, 1, dt12.Rows.Count - 1);
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
                    Table oTbl평가방법 = rUtil.GetTable(lstTable, "@B11DmobNm@");
                    Table oTbl총괄표 = rUtil.GetTable(lstTable, "@B12ExpsGrpNm@");

                    rUtil.ReplaceTextAllParagraph(doc, "@B7Title@", pdr7["Vitm"] + "");

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
                                rUtil.ReplaceTableRow(oTbl평가방법.GetRow(i + 0), sKey, sValue);
                            }
                        }
                    }

                    double ReqSum = 0;
                    double DoLosSum = 0;

                    dtB = dt12;
                    sPrefix = "B12";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        rUtil.TableMergeCellsV(oTbl총괄표, 0, 1, dtB.Rows.Count + 1);
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            string expscd = dr["ExpsCd"] + "";
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "ReqAmt")
                                {
                                    if (expscd != "300270001") ReqSum += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "DoLosAmt")
                                {
                                    if (expscd != "300270001") DoLosSum += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTbl총괄표.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTable(oTbl총괄표, "@db12ReqSum@", Utils.AddComma(ReqSum));
                    rUtil.ReplaceTable(oTbl총괄표, "@db12DoLosSum@", Utils.AddComma(DoLosSum));
                    rUtil.ReplaceTable(oTbl총괄표, "@B7Vitm@", pdr7["Vitm"] + "");

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
