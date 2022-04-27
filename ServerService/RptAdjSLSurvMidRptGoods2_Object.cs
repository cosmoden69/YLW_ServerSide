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
    public class RptAdjSLSurvMidRptGoods2_Object
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
                    Table oTbl평가기준 = rUtil.GetTable(lstTable, "1) 평가기준");

                    //평가결과 행추가
                    drs = pds.Tables["DataBlock16"]?.Select("1 = 1");
                    if (drs != null)
                    {
                        sKey = rUtil.GetFieldName("B16", "ObjDvs");
                        Table oTableA = rUtil.GetSubTable(oTbl평가기준, sKey);
                        if (oTableA != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRow(oTableA, 1, drs.Length - 1);
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
                    Table oTbl평가기준 = rUtil.GetTable(lstTable, "1) 평가기준");

                    //변수가 replace 되기 전에 테이블을 찾아 놓는다
                    sKey = rUtil.GetFieldName("B16", "ObjDvs");
                    Table oTableA = rUtil.GetSubTable(oTbl평가기준, sKey);

                    dtB = pds.Tables["DataBlock3"];
                    sPrefix = "B3";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        DataRow dr = dtB.Rows[0];

                        foreach (DataColumn col in dtB.Columns)
                        {
                            sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                            sValue = dr[col] + "";
                            if (col.ColumnName == "NglgSetoffAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "LosAmt") sValue = Utils.AddComma(sValue);
                            if (col.ColumnName == "AgrmAmt") sValue = Utils.AddComma(sValue);
                            rUtil.ReplaceTextAllParagraph(doc, sKey, sValue);
                            rUtil.ReplaceTables(lstTable, sKey, sValue);
                        }
                    }

                    double db16InsDmndAmt = 0;
                    double db16ObjCstrAmt = 0;

                    dtB = pds.Tables["DataBlock16"];
                    sPrefix = "B16";
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
                                if (col.ColumnName == "InsDmndAmt")
                                {
                                    db16InsDmndAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                if (col.ColumnName == "ObjCstrAmt")
                                {
                                    db16ObjCstrAmt += Utils.ToDouble(sValue);
                                    sValue = Utils.AddComma(sValue);
                                }
                                rUtil.ReplaceTableRow(oTableA.GetRow(i + 1), sKey, sValue);
                            }
                        }
                    }
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db16InsDmndAmt@", Utils.AddComma(db16InsDmndAmt));
                    rUtil.ReplaceTableRow(oTableA.GetRow(dtB.Rows.Count + 1), "@db16ObjCstrAmt@", Utils.AddComma(db16ObjCstrAmt));

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
