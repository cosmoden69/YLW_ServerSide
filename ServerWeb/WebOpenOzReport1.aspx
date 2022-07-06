<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebOpenOzReport1.aspx.cs" Inherits="YLW_WebService.ServerSide.WebOpenOzReport1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>OZReport</title>
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <!--20210914.tjjang, ACE인 경우에는 /Oz70 를 넣어야 하고, EVER 인 경우 /Oz70 을 뺴야한다-->
    <script type="text/javascript" src="http://ksystem.metro070.com:8200/Oz70/ozhviewer/jquery-1.8.3.min.js"></script>
    <link rel="stylesheet" href="http://ksystem.metro070.com:8200/Oz70/ozhviewer/jquery-ui.css" type="text/css" />
    <script type="text/javascript" src="http://ksystem.metro070.com:8200/Oz70/ozhviewer/jquery-ui.min.js"></script>
    <link rel="stylesheet" href="http://ksystem.metro070.com:8200/Oz70/ozhviewer/ui.dynatree.css?WEBK20190919=12" type="text/css" />
    <script type="text/javascript" src="http://ksystem.metro070.com:8200/Oz70/ozhviewer/jquery.dynatree.js?WEBK20190919=12" charset="utf-8"></script>
    <script type="text/javascript" src="http://ksystem.metro070.com:8200/Oz70/ozhviewer/OZJSViewer.js?WEBK20190919=12" charset="utf-8"></script>
</head>
<body style="height:100%">
    <div id="OZViewer" style="width:100%; height:100%;"></div>
    <script type="text/javascript">
    	
			function SetOZParamters_OZViewer(){
				var oz;
				var strOdiFileName;
				var strSiteFileName;
				var xmlString;
				oz = document.getElementById("OZViewer");
                  //Default -> 레포트설정에 따라 변경가능
                oz.sendToActionScript("viewer.isframe", "true");
                  //print미리보기 제거
                oz.sendToActionScript("print.mode", "false");
                  //최소 글꼴 크기 조정
                oz.sendToActionScript("viewer.fontdpi", "auto");
                  //#20200806 서브폼 출력물의 경우 100%를 기본으로 설정
                oz.sendToActionScript("viewer.zoom", "100");
	      
				strOdiFileName = "RptAdjSLInvoiceViewBillIssueIn";
				strSiteFileName = "RptAdjSLInvoiceViewBillIssueIn";
				xmlString = "xmlData=<?xml version='1.0' encoding='UTF-8'?><ROOT><DataBlock13><CustName>MG손해보험(주)</CustName><InsurDept>장기보상지원파트 장기재물보상센터</InsurDept><InsurChrg>최성진</InsurChrg><InsurChrgMail>jin@mggenins.com</InsurChrgMail><LasRprtNo>AA  22040001</LasRprtNo><AcptDt>2021년 02월 04일</AcptDt><SurvAsgnEmpSeq>0</SurvAsgnEmpSeq><SurvAsgnEmpName></SurvAsgnEmpName><DcmgDt></DcmgDt><CclsDt>2022년 06월 29일</CclsDt><LasRptSbmsDt>2022년 06월 29일</LasRptSbmsDt><InsurPrdt></InsurPrdt><Insurant></Insurant><InsurNo></InsurNo><Insured></Insured><AcdtNo></AcdtNo><AcdtDt></AcdtDt><Vitm></Vitm><InvcAdjFeeRmk>ㄱㄱ</InvcAdjFeeRmk><InvcAdjFee>280000.00000</InvcAdjFee><InvcIctvRmk>ㄴㄴ</InvcIctvRmk><InvcIctvAmt>1.00000</InvcIctvAmt><DayExpsRmk></DayExpsRmk><DayExps></DayExps><TrspExpsRmk></TrspExpsRmk><TrspExps></TrspExps><OilAmt></OilAmt><DocuRmk></DocuRmk><DocuAmt></DocuAmt><InvcCsltReqRmk>ㅁ</InvcCsltReqRmk><InvcCsltReqAmt>0.00000</InvcCsltReqAmt><OthRmk></OthRmk><OthAmt></OthAmt><DmndSubTot>280001.00000</DmndSubTot><GroupKey>1/393170/1/1</GroupKey><CostRcptFileSeq>25242</CostRcptFileSeq></DataBlock13><DataBlock14><PrgMgtCdName>11.수임</PrgMgtCdName><PrgMgtDt>2022년 02월 04일</PrgMgtDt><JobCnts></JobCnts><DtlTrtCnts></DtlTrtCnts><GroupKey>1/393170/1/1</GroupKey></DataBlock14><DataBlock14><PrgMgtCdName>12.조사자배당</PrgMgtCdName><PrgMgtDt>2022년 02월 04일</PrgMgtDt><JobCnts></JobCnts><DtlTrtCnts></DtlTrtCnts><GroupKey>1/393170/1/1</GroupKey></DataBlock14><DataBlock15><BsnTrpFrDt></BsnTrpFrDt><SurvAsgnEmpSeqName></SurvAsgnEmpSeqName><DayExps></DayExps><OilAmt></OilAmt><TrspExps></TrspExps><DocuAmt></DocuAmt><OthAmt></OthAmt><DmndSubTot></DmndSubTot><DptrRign>출장지 :</DptrRign><TrspExpsRmk>교통비내역 :</TrspExpsRmk><GroupKey>1/393170/1/1</GroupKey></DataBlock15><DataBlock16><DeptHeadCdName>인보험</DeptHeadCdName><InsurDvsCdName>질병</InsurDvsCdName><InsurSubTypCdName></InsurSubTypCdName><DyasGap>510</DyasGap><InsurGivTypCdName></InsurGivTypCdName><InsurKeepCdName></InsurKeepCdName><CloseName></CloseName><InsurDmndAmt>0.00000</InsurDmndAmt><InsurGivAmt></InsurGivAmt><SaveAmt>0.00000</SaveAmt><GroupKey>1/393170/1/1</GroupKey></DataBlock16></ROOT>";
				oz.sendToActionScript("odi.odinames", strOdiFileName); //"xml");
				oz.sendToActionScript("odi." + strOdiFileName + ".usescheduleddata", "ozp:///sdmmaker_html5(xml)/sdmmaker_ylw_h.js");
				oz.sendToActionScript("odi." + strOdiFileName + ".pcount", "1");
				oz.sendToActionScript("odi." + strOdiFileName + ".args1", xmlString);
				
				oz.sendToActionScript("connection.servlet", "http://ksystem.metro070.com:8200/Oz70/");
				oz.sendToActionScript("connection.reportname", "/" + strSiteFileName + ".ozr");
				oz.sendToActionScript("connection.displayname", strOdiFileName + "_pdf");
				oz.sendToActionScript("connection.formfromserver", "true");
				
				//폰트 파라미터
				oz.sendToActionScript("pdf.fontembedding", "true");
				oz.sendToActionScript("information.debug", "true");
				
				//SDM FILE 에러 수정적용
				oz.sendToActionScript("connection.datafromserver", "false");
				return true;
			}
			
			var opt = [];
			opt["print_exportfrom"] = "scheduler";
        start_ozjs("OZViewer", "http://ksystem.metro070.com:8200/Oz70/ozhviewer/", opt);
    </script>
</body>
</html>
