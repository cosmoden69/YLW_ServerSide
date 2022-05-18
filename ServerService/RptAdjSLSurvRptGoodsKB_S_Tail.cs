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
    public class RptAdjSLSurvRptGoodsKB_S_Tail
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile)
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
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl유첨서류 = rUtil.GetTable(lstTable, "@B9FileNo@");

                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl당사, 1, drs.Length - 1);
                        }
                    }
                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl옥션, 1, drs.Length - 1);
                        }
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
                    if (dtB != null)
                    {
                        if (oTbl유첨서류 != null)
                        {
                            //테이블의 끝에 추가
                            rUtil.TableAddRow(oTbl유첨서류, 1, dtB.Rows.Count - 1);
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
                    Table oTbl당사 = rUtil.GetTable(lstTable, "@B8RmnObjCost@");
                    Table oTbl옥션 = rUtil.GetTable(lstTable, "@B8SucBidDt@");
                    Table oTbl유첨서류 = rUtil.GetTable(lstTable, "@B9FileNo@");


                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 1");
                    sPrefix = "B8";
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl당사 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    rUtil.ReplaceTableRow(oTbl당사.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl당사.Remove();
                    }

                    drs = pds.Tables["DataBlock8"]?.Select("TrtCd % 10 = 2");
                    sPrefix = "B8";
                    if (drs != null && drs.Length > 0)
                    {
                        if (oTbl옥션 != null)
                        {
                            for (int i = 0; i < drs.Length; i++)
                            {
                                DataRow dr = drs[i];
                                foreach (DataColumn col in dr.Table.Columns)
                                {
                                    sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                    sValue = dr[col] + "";
                                    if (col.ColumnName == "RmnObjCnt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjCost") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "RmnObjAmt") sValue = Utils.AddComma(sValue);
                                    if (col.ColumnName == "AuctFrDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "AuctToDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    if (col.ColumnName == "SucBidDt") sValue = Utils.DateFormat(sValue, "yyyy.MM.dd");
                                    rUtil.ReplaceTableRow(oTbl옥션.GetRow(i + 1), sKey, sValue);
                                }
                            }
                        }
                    }
                    else
                    {
                        oTbl옥션.Remove();
                    }

                    dtB = pds.Tables["DataBlock9"];
                    sPrefix = "B9";
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
                                    if (col.ColumnName == "FileAmt") sValue = Utils.AddComma(sValue == "" || sValue == "0" ? "1" : sValue) + "부";
                                    rUtil.ReplaceTableRow(oTable.GetRow(i + 1), sKey, sValue);
                                }
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
