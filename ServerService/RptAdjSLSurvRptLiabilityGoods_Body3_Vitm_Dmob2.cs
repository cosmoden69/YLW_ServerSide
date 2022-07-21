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
    public class RptAdjSLSurvRptLiabilityGoods_Body3_Vitm_Dmob2
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile, DataRow pdr11, string num)
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
                string vitmSubSeq = Utils.ConvertToString(pdr11["VitmSubSeq"]);
                string dmobSeq = Utils.ConvertToString(pdr11["DmobSeq"]);
                drs = pds.Tables["DataBlock15"]?.Select("VitmSubSeq = " + vitmSubSeq + " AND DmobSeq = " + dmobSeq + " ", " InsurObjSerl ASC");
                DataTable dt15 = (drs.Length < 1 ? pds.Tables["DataBlock15"].Clone() : drs.CopyToDataTable());

                System.IO.File.Copy(sDocFile, sWriteFile, true);

                using (WordprocessingDocument wDoc = WordprocessingDocument.Open(sWriteFile, true))
                {
                    MainDocumentPart mDoc = wDoc.MainDocumentPart;
                    Document doc = mDoc.Document;
                    RptUtils rUtil = new RptUtils(mDoc);

                    IEnumerable<Table> lstTable = doc.Body.Elements<Table>();
                    Table oTbl평가결과 = rUtil.GetTable(lstTable, "@B15InsurObjSerl@");

                    if (dt15 != null)
                    {
                        if (oTbl평가결과 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl평가결과, 2, 2, dt15.Rows.Count - 1);
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
                    Table oTbl평가결과 = rUtil.GetTable(lstTable, "@B15InsurObjSerl@");

                    rUtil.ReplaceTextAllParagraph(doc, "@B11Num@", num);
                    rUtil.ReplaceTextAllParagraph(doc, "@B11DmobNm@", pdr11["DmobNm"] + "");

                    dtB = dt15;
                    sPrefix = "B15";
                    if (dtB != null)
                    {
                        if (dtB.Rows.Count < 1) dtB.Rows.Add();
                        for (int i = 0; i < dtB.Rows.Count; i++)
                        {
                            DataRow dr = dtB.Rows[i];
                            int rnum = (i + 1) * 2;
                            foreach (DataColumn col in dtB.Columns)
                            {
                                sKey = rUtil.GetFieldName(sPrefix, col.ColumnName);
                                sValue = dr[col] + "";
                                if (col.ColumnName == "ObjArea") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "ObjCost") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "ObjDprcTotRate") sValue = string.Format("{0:0.00}", double.Parse(sValue));
                                if (col.ColumnName == "ObjRePurcAmt") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "ObjInsureValue") sValue = Utils.AddComma(sValue);
                                if (col.ColumnName == "LosAmt") sValue = Utils.AddComma(sValue);
                                rUtil.ReplaceTableRow(oTbl평가결과.GetRow(rnum + 0), sKey, sValue);
                                rUtil.ReplaceTableRow(oTbl평가결과.GetRow(rnum + 1), sKey, sValue);
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
