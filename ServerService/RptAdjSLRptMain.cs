using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using System.Xml.Schema;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

using YLWService;

namespace YLW_WebService.ServerSide
{
    public class RptAdjSLRptMain
    {//인보험
        public string myPath = Application.StartupPath;

        public Response RptMain(ReportParam para, ref string rptPath, ref string rptName)
        {
            Alert.WriteHist("RptMain", JsonConvert.SerializeObject(para));

            if (para.ReportName == "SaveReport")   //저장된 리포트를 반환
            {
                RptAdjSLXSaveReport rpt = new RptAdjSLXSaveReport(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            else if (para.ReportName == "SaveReportPers")   //저장된 리포트를 반환(인보험)
            {
                RptAdjSLXSaveReport rpt = new RptAdjSLXSaveReport(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }

            //종결보고서
            //==============================================================================================
            //출력설계_1511_서식_종결보고서_표준(인)
            if (para.ReportName == "RptAdjSLRptSurvRptPers")
            {
                RptAdjSLRptSurvRptPers rpt = new RptAdjSLRptSurvRptPers(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1533_서식_종결보고서(인보험_메리츠)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMeritz")
            {
                RptAdjSLRptSurvRptPersMeritz rpt = new RptAdjSLRptSurvRptPersMeritz(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1534_서식_종결보고서(인보험_현대해상)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersHyundai")
            {
                RptAdjSLRptSurvRptPersHyundai rpt = new RptAdjSLRptSurvRptPersHyundai(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1535_서식_종결보고서(인보험_흥국화재)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersHeungkuk")
            {
                RptAdjSLRptSurvRptPersHeungkuk rpt = new RptAdjSLRptSurvRptPersHeungkuk(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1536_서식_종결보고서(인보험_DB생명)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDBLife")
            {
                RptAdjSLRptSurvRptPersDBLife rpt = new RptAdjSLRptSurvRptPersDBLife(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1537_서식_종결보고서(인보험_DB손해)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDBLoss")
            {
                RptAdjSLRptSurvRptPersDBLoss rpt = new RptAdjSLRptSurvRptPersDBLoss(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1538_서식_종결보고서(인보험_MG손해_단순)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMGLossSmpl")
            {
                RptAdjSLRptSurvRptPersMGLossSmpl rpt = new RptAdjSLRptSurvRptPersMGLossSmpl(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1539_서식_종결보고서(인보험_MG손해_일반)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMGLoss")
            {
                RptAdjSLRptSurvRptPersMGLoss rpt = new RptAdjSLRptSurvRptPersMGLoss(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //라이나 전문보고서 출력
            else if (para.ReportName == "DlgAdjSLSurvRptLina")
            {
                DlgAdjSLSurvRptLina rpt = new DlgAdjSLSurvRptLina(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //KDB 전문보고서 출력 (보고서 설계 없음)
            else if (para.ReportName == "DlgAdjSLSurvRptKDB")
            {
                //DlgAdjSLSurvRptKDB rpt = new DlgAdjSLSurvRptKDB(myPath);
                //return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //삼성계약적부 전문보고서 출력
            else if (para.ReportName == "RptAdjSLRptEDISSL")
            {
                RptAdjSLRptEDISSL rpt = new RptAdjSLRptEDISSL(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2551_서식_종결보고서(재물)
            else if (para.ReportName == "RptAdjSLSurvRptGoods")
            {
                RptAdjSLSurvRptGoods rpt = new RptAdjSLSurvRptGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2552_서식_종결보고서(재물-대물)
            else if (para.ReportName == "RptAdjSLSurvRptGoods2")
            {
                RptAdjSLSurvRptGoods2 rpt = new RptAdjSLSurvRptGoods2(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2553_서식_종결보고서(배책-대인)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityPers")
            {
                RptAdjSLSurvRptLiabilityPers rpt = new RptAdjSLSurvRptLiabilityPers(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2554_서식_종결보고서(배책-대물)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityGoods")
            {
                RptAdjSLSurvRptLiabilityGoods rpt = new RptAdjSLSurvRptLiabilityGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //출력설계_2561_서식_농협_종결보고서(재물)
            else if (para.ReportName == "RptAdjSLSurvRptGoodsNH")
            {
                // 2021-07-14 수정(폰트등) AcptMgmtSeq = 392272
                RptAdjSLSurvRptGoodsNH rpt = new RptAdjSLSurvRptGoodsNH(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2562_서식_농협_종결보고서(재물-대물, 배책-차량)
            else if (para.ReportName == "RptAdjSLSurvRptGoods2NH")
            {
                RptAdjSLSurvRptLiabilityGoodsNH rpt = new RptAdjSLSurvRptLiabilityGoodsNH(myPath);  //배책-대물(차량)과 동일
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2571_서식_농협_종결보고서(배책-대인)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityPersNH")
            {
                RptAdjSLSurvRptLiabilityPersNH rpt = new RptAdjSLSurvRptLiabilityPersNH(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2562_서식_농협_종결보고서(재물-대물, 배책-차량)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityGoodsNH")
            {
                RptAdjSLSurvRptLiabilityGoodsNH rpt = new RptAdjSLSurvRptLiabilityGoodsNH(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //출력설계_2563_서식_농협_종결보고서(재물, 간편)
            else if (para.ReportName == "RptAdjSLSurvRptGoodsNH_S")
            {
                RptAdjSLSurvRptGoodsNH_S rpt = new RptAdjSLSurvRptGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2564_서식_농협_종결보고서(재물-대물, 간편)
            else if (para.ReportName == "RptAdjSLSurvRptGoods2NH_S")
            {
                RptAdjSLSurvRptGoods2NH_S rpt = new RptAdjSLSurvRptGoods2NH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2573_서식_농협_종결보고서(배책-대인, 간편)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityPersNH_S")
            {
                RptAdjSLSurvRptLiabilityPersNH_S rpt = new RptAdjSLSurvRptLiabilityPersNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2574_서식_농협_종결보고서(배책-차량, 간편)
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityGoodsNH_S")
            {
                RptAdjSLSurvRptLiabilityGoodsNH_S rpt = new RptAdjSLSurvRptLiabilityGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //종결보고서(재물 - DB) : 2581
            else if (para.ReportName == "RptAdjSLSurvRptGoodsDB")
            {
                RptAdjSLSurvRptGoodsDB rpt = new RptAdjSLSurvRptGoodsDB(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //종결보고서(배책 - DB) : 2582
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityDB")
            {
                RptAdjSLSurvRptLiabilityDB rpt = new RptAdjSLSurvRptLiabilityDB(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //종결보고서(재물, 간편 - KB) : 2591
            else if (para.ReportName == "RptAdjSLSurvRptGoodsKB_S")
            {
                RptAdjSLSurvRptGoodsKB_S rpt = new RptAdjSLSurvRptGoodsKB_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //종결보고서(재물-대물, 간편 - KB) : 2592
            else if (para.ReportName == "RptAdjSLSurvRptGoods2KB_S")
            {
                RptAdjSLSurvRptGoods2KB_S rpt = new RptAdjSLSurvRptGoods2KB_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //종결보고서(배책-대인) : 2593
            else if (para.ReportName == "RptAdjSLSurvRptLiabilityPersKB")
            {
                RptAdjSLSurvRptLiabilityPersKB rpt = new RptAdjSLSurvRptLiabilityPersKB(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //중간보고서
            //==============================================================================================
            //출력설계_1511_서식_중간보고서_표준(인)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMid")
            {
                RptAdjSLRptSurvRptPersMid rpt = new RptAdjSLRptSurvRptPersMid(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1533_서식_종결보고서(인보험_메리츠)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMeritzMid")
            {
                RptAdjSLRptSurvRptPersMeritz rpt = new RptAdjSLRptSurvRptPersMeritz(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1534_서식_종결보고서(인보험_현대해상)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersHyundaiMid")
            {
                RptAdjSLRptSurvRptPersHyundai rpt = new RptAdjSLRptSurvRptPersHyundai(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1535_서식_종결보고서(인보험_흥국화재)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersHeungkukMid")
            {
                RptAdjSLRptSurvRptPersHeungkuk rpt = new RptAdjSLRptSurvRptPersHeungkuk(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1536_서식_종결보고서(인보험_DB생명)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDBLifeMid")
            {
                RptAdjSLRptSurvRptPersDBLife rpt = new RptAdjSLRptSurvRptPersDBLife(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1537_서식_종결보고서(인보험_DB손해)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDBLossMid")
            {
                RptAdjSLRptSurvRptPersDBLoss rpt = new RptAdjSLRptSurvRptPersDBLoss(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1538_서식_종결보고서(인보험_MG손해_단순)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMGLossSmplMid")
            {
                RptAdjSLRptSurvRptPersMGLossSmpl rpt = new RptAdjSLRptSurvRptPersMGLossSmpl(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1539_서식_종결보고서(인보험_MG손해_일반)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersMGLossMid")
            {
                RptAdjSLRptSurvRptPersMGLoss rpt = new RptAdjSLRptSurvRptPersMGLoss(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2531_서식_중간보고서(재물)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoods")
            {
                RptAdjSLSurvMidRptGoods rpt = new RptAdjSLSurvMidRptGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2532_서식_중간보고서(재물-대물)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoods2")
            {
                RptAdjSLSurvMidRptGoods2 rpt = new RptAdjSLSurvMidRptGoods2(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2533_서식_중간보고서(배책-대인)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityPers")
            {
                RptAdjSLSurvMidRptLiabilityPers rpt = new RptAdjSLSurvMidRptLiabilityPers(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2534_서식_중간보고서(배책-대물)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityGoods")
            {
                RptAdjSLSurvMidRptLiabilityGoods rpt = new RptAdjSLSurvMidRptLiabilityGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //출력설계_2565_서식_농협_진행보고서(재물)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoodsNH")
            {
                RptAdjSLSurvMidRptGoodsNH rpt = new RptAdjSLSurvMidRptGoodsNH(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2567_서식_농협_진행보고서(재물-대물)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoods2NH")
            {
                RptAdjSLSurvMidRptGoods2NH_S rpt = new RptAdjSLSurvMidRptGoods2NH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2575_서식_농협_진행보고서(배책-대인)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityPersNH")
            {
                RptAdjSLSurvMidRptLiabilityPersNH_S rpt = new RptAdjSLSurvMidRptLiabilityPersNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2576_서식_농협_진행보고서(배책-차량)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityGoodsNH")
            {
                RptAdjSLSurvMidRptLiabilityGoodsNH_S rpt = new RptAdjSLSurvMidRptLiabilityGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //출력설계_2566_서식_농협_진행보고서(재물, 간편)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoodsNH_S")
            {
                RptAdjSLSurvMidRptGoodsNH_S rpt = new RptAdjSLSurvMidRptGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2567_서식_농협_진행보고서(재물-대물, 간편)
            else if (para.ReportName == "RptAdjSLSurvMidRptGoods2NH_S")
            {
                RptAdjSLSurvMidRptGoods2NH_S rpt = new RptAdjSLSurvMidRptGoods2NH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2575_서식_농협_진행보고서(배책-대인, 간편)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityPersNH_S")
            {
                RptAdjSLSurvMidRptLiabilityPersNH_S rpt = new RptAdjSLSurvMidRptLiabilityPersNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2576_서식_농협_진행보고서(배책-차량, 간편)
            else if (para.ReportName == "RptAdjSLSurvMidRptLiabilityGoodsNH_S")
            {
                RptAdjSLSurvMidRptLiabilityGoodsNH_S rpt = new RptAdjSLSurvMidRptLiabilityGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //현장보고서
            //==============================================================================================
            //출력설계_2511_서식_현장보고서(재물)
            else if (para.ReportName == "RptAdjSLSurvSpotRptGoods")
            {
                RptAdjSLSurvSpotRptGoods rpt = new RptAdjSLSurvSpotRptGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2512_서식_현장보고서(재물-대물)
            else if (para.ReportName == "RptAdjSLSurvSpotRptGoods2")
            {
                RptAdjSLSurvSpotRptGoods2 rpt = new RptAdjSLSurvSpotRptGoods2(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //농협_현장보고서(재물, 간편) -- 2021-2-23 현재 출력설계서 없음
            else if (para.ReportName == "RptAdjSLSurvSpotRptGoodsNH_S")
            {
                //RptAdjSLSurvSpotRptGoodsNH_S rpt = new RptAdjSLSurvSpotRptGoodsNH_S(myPath);
                //return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2568_서식_농협_현장보고서(재물-대물, 간편)
            else if (para.ReportName == "RptAdjSLSurvSpotRptGoods2NH_S")
            {
                RptAdjSLSurvSpotRptGoods2NH_S rpt = new RptAdjSLSurvSpotRptGoods2NH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2577_서식_농협_현장보고서(배책-대인, 간편)
            else if (para.ReportName == "RptAdjSLSurvSpotRptLiabilityPersNH_S")
            {
                RptAdjSLSurvSpotRptLiabilityPersNH_S rpt = new RptAdjSLSurvSpotRptLiabilityPersNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2578_서식_농협_현장보고서(배책-차량, 간편)
            else if (para.ReportName == "RptAdjSLSurvSpotRptLiabilityGoodsNH_S")
            {
                RptAdjSLSurvSpotRptLiabilityGoodsNH_S rpt = new RptAdjSLSurvSpotRptLiabilityGoodsNH_S(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //----------------------------------------------------------------------------------------------
            //현장보고서(재물 - DB) : 2585
            else if (para.ReportName == "RptAdjSLSurvSpotRptGoodsDB")
            {
                RptAdjSLSurvSpotRptGoodsDB rpt = new RptAdjSLSurvSpotRptGoodsDB(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //현장보고서(배책 - DB) : 2586
            else if (para.ReportName == "RptAdjSLSurvSpotRptLiabilityDB")
            {
                RptAdjSLSurvSpotRptLiabilityDB rpt = new RptAdjSLSurvSpotRptLiabilityDB(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //손해사정서
            //==============================================================================================
            //출력설계_1522_서식_손해사정서(인보험)
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDmg")
            {
                RptAdjSLRptSurvRptPersDmg rpt = new RptAdjSLRptSurvRptPersDmg(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2541_서식_손해사정서(재물)
            else if (para.ReportName == "RptAdjSLSurvDmgRptGoods")
            {
                RptAdjSLSurvDmgRptGoods rpt = new RptAdjSLSurvDmgRptGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2542_서식_손해사정서(재물-대물)
            else if (para.ReportName == "RptAdjSLSurvDmgRptGoods2")
            {
                RptAdjSLSurvDmgRptGoods2 rpt = new RptAdjSLSurvDmgRptGoods2(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2631_서식_교부용 손해사정서_재물_표준
            else if (para.ReportName == "RptAdjSLSurvDmgRptGoodsN")
            {
                RptAdjSLSurvDmgRptGoodsN rpt = new RptAdjSLSurvDmgRptGoodsN(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2632_서식_교부용 손해사정서_재물_현대해상
            else if (para.ReportName == "RptAdjSLSurvDmgRptGoodsN_HD")
            {
                RptAdjSLSurvDmgRptGoodsN_HD rpt = new RptAdjSLSurvDmgRptGoodsN_HD(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2641_서식_교부용 손해사정서_배상_표준
            else if (para.ReportName == "RptAdjSLSurvDmgRptLiabilityGoods")
            {
                RptAdjSLSurvDmgRptLiabilityGoods rpt = new RptAdjSLSurvDmgRptLiabilityGoods(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_2642_서식_교부용 손해사정서_배상_현대해상
            else if (para.ReportName == "RptAdjSLSurvDmgRptLiabilityGoods_HD")
            {
                RptAdjSLSurvDmgRptLiabilityGoods_HD rpt = new RptAdjSLSurvDmgRptLiabilityGoods_HD(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //손해사정서 교부동의서 (인보험)
            //==============================================================================================
            //출력설계_1525_서식_손해사정서 교부 동의 및 확인서_공통
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDmg1525")
            {
                RptAdjSLRptSurvRptPersDmg1525 rpt = new RptAdjSLRptSurvRptPersDmg1525(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1526_서식_손해사정서 교부 동의 및 확인서_라이나생명
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDmg1526")
            {
                RptAdjSLRptSurvRptPersDmg1526 rpt = new RptAdjSLRptSurvRptPersDmg1526(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1527_서식_손해사정서 교부 동의 및 확인서_삼성화재
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDmg1527")
            {
                RptAdjSLRptSurvRptPersDmg1527 rpt = new RptAdjSLRptSurvRptPersDmg1527(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //출력설계_1528_서식_손해사정서 교부 동의 및 확인서_흥국화재
            else if (para.ReportName == "RptAdjSLRptSurvRptPersDmg1528")
            {
                RptAdjSLRptSurvRptPersDmg1528 rpt = new RptAdjSLRptSurvRptPersDmg1528(myPath);
                return rpt.GetReport(para, ref rptPath, ref rptName);
            }
            //==============================================================================================
            return new Response() { Result = -1, Message = "해당하는 보고서가 없습니다" };
        }
    }
}