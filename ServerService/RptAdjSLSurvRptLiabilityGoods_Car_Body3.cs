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
    public class RptAdjSLSurvRptLiabilityGoods_Car_Body3
    {
        public string SetSample1(string sDocFile, string sXSDFile, DataSet pds, string sWriteFile, int ObjSeq)
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B7DmobSortNo@");
                    Table oTbl피해물파손부위확인 = rUtil.GetTable(lstTable, "@B8DmobSortNo@");

                    dtB = pds.Tables["DataBlock7"];
                    drs = dtB?.Select("ObjSeq = " + ObjSeq); //목적물 별로 RePlace
                    var B7RowCnt = drs.Length;
                    if (dtB != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            //테이블의 끝에 추가
                            double cnt = Math.Truncate((B7RowCnt + 1) / 2.0);
                            for (int i = 1; i < cnt; i++)
                            {
                                rUtil.TableAddRow(oTbl손해상황, 1, 1);
                                rUtil.TableAddRow(oTbl손해상황, 2, 1);
                            }
                        }
                    }

                    
                    dtB = pds.Tables["DataBlock8"];
                    drs = dtB?.Select("ObjSeq = " + ObjSeq); //목적물 별로 RePlace
                    var B8RowCnt = drs.Length;
                    if (dtB != null)
                    {
                        //피해물 파손부위확인
                        if (oTbl피해물파손부위확인 != null)
                        {
                            //테이블의 중간에 삽입
                            rUtil.TableInsertRows(oTbl피해물파손부위확인, 1, 2, B8RowCnt - 1);
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
                    Table oTbl손해상황 = rUtil.GetTable(lstTable, "@B7DmobSortNo@");
                    Table oTbl피해물파손부위확인 = rUtil.GetTable(lstTable, "@B8DmobSortNo@");
                    
                    dtB = pds.Tables["DataBlock7"];
                    sPrefix = "B7";
                    drs = dtB?.Select("ObjSeq = " + ObjSeq);
                    List<DataRow> list7 = drs.ToList();

                    if (drs != null)
                    {
                        if (oTbl손해상황 != null)
                        {
                            //if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock7"].Rows.Add(0) };
                            if (drs.Length < 1) //하나라도 없을 경우 기본 1행 추가
                            {
                                list7.Add(pds.Tables["DataBlock7"].Rows.Add(0));
                                drs = list7.ToArray();
                            }
                            if (drs.Length % 2 == 1) //홀수인 경우 1행 추가
                            {
                                list7.Add(pds.Tables["DataBlock7"].Rows.Add(0));
                                drs = list7.ToArray();
                            }
                            
                            for (int i = 0; i < drs.Length; i++)
                            {
                                int rnum = (int)Math.Truncate(i / 2.0) * 2 + 1;
                                int rmdr = i % 2;
                                
                                
                                //순번
                                sKey = rUtil.GetFieldName(sPrefix, "DmobSortNo");
                                sValue = Utils.ConvertToString(drs[i]["DmobSortNo"] + "");
                                TableRow xrow0 = oTbl손해상황.GetRow(0);
                                rUtil.SetText(xrow0.GetCell(0), sKey, sValue);
                                
                                
                                //이미지
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                //sValue = dr["AcdtPictImage"] + "";
                                sValue = Utils.ConvertToString(drs[i]["AcdtPictImage"] + "");
                                TableRow xrow1 = oTbl손해상황.GetRow(rnum);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 2700000L, 2000000L);
                                }
                                catch { }
                                
                                
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = Utils.ConvertToString(drs[i]["AcdtPictCnts"] + "");
                                TableRow xrow2 = oTbl손해상황.GetRow(rnum + 1);
                                rUtil.SetText(xrow2.GetCell(rmdr), sKey, sValue);
                                
                            }
                        }
                    }
        


                    dtB = pds.Tables["DataBlock8"];
                    sPrefix = "B8";
                    drs = dtB?.Select("ObjSeq = " + ObjSeq);
                    List<DataRow> list8 = drs.ToList();
                    if (drs != null)
                    {
                        if (oTbl피해물파손부위확인 != null)
                        {

                            //if (drs.Length < 1) drs = new DataRow[1] { pds.Tables["DataBlock8"].Rows.Add() };
                            
                            if (drs.Length < 1) //하나라도 없을 경우 기본 1행 추가
                            {
                                list8.Add(pds.Tables["DataBlock7"].Rows.Add(0));
                                drs = list8.ToArray();
                            }
                            for (int i = 0; i < drs.Length; i++)
                            {
                                int rnum = (int)Math.Truncate(i / 1.0) * 2;
                                int rmdr = i % 1;

                                //순번
                                sKey = rUtil.GetFieldName(sPrefix, "DmobSortNo");
                                sValue = Utils.ConvertToString(drs[i]["DmobSortNo"] + "");
                                TableRow xrow0 = oTbl피해물파손부위확인.GetRow(rnum);
                                rUtil.SetText(xrow0.GetCell(rmdr), sKey, sValue);

                                //이미지
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictImage");
                                //sValue = dr["AcdtPictImage"] + "";
                                sValue = Utils.ConvertToString(drs[i]["AcdtPictImage"] + "");

                                TableRow xrow1 = oTbl피해물파손부위확인.GetRow(rnum + 1);
                                rUtil.SetText(xrow1.GetCell(rmdr), sKey, "");
                                try
                                {
                                    Image img = Utils.stringToImage(sValue);
                                    rUtil.SetImageNull(xrow1.GetCell(rmdr), img, 50000L, 50000L, 6200000L, 4000000L);
                                }
                                catch { }

                                //설명
                                sKey = rUtil.GetFieldName(sPrefix, "AcdtPictCnts");
                                sValue = Utils.ConvertToString(drs[i]["AcdtPictCnts"] + "");
                                TableRow xrow2 = oTbl피해물파손부위확인.GetRow(rnum + 2);
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
