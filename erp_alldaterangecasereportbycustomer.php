<?php
  session_start();
  $pagetitle = "業務部 &raquo; 期間內帳款客到貨合計";
  include("_data.php");
  include("_erp.php");
 // auth("erp_alldaterangecasereportbycustomer.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }     
  
  if (is_null($_GET['edate'])) {
    $edate =  date('Y-m-d');
  } else {
    $edate=$_GET['edate'];
  }                                
  if ($_GET["submit"]=="匯出") {  
      if ($_GET['clienttype']=='1') {
          $clientfilter=' and oea17=occ01 ';       
      } else {
          $clientfilter=' and oea04=occ01 ';   
      }
    
      error_reporting(E_NONE);  
      require_once 'classes/PHPExcel.php'; 
      require_once 'classes/PHPExcel/IOFactory.php';  
      $objReader = PHPExcel_IOFactory::createReader('Excel5');
      $objPHPExcel = $objReader->load("templates/erp_daterangecasesbycustomer.xls");    
      
      $s2="select occ02, 
        sum(a93111) a93111, sum(a93112) a93112, sum(a93113) a93113, sum(a93121) a93121, sum(a93122) a93122, sum(a93123) a93123,  
        sum(a93131) a93131, sum(a93132) a93132, sum(a93133) a93133, sum(a93141) a93141, sum(a93142) a93142, sum(a93143) a93143,  
        sum(a93151) a93151, sum(a93152) a93152, sum(a93153) a93153, sum(a93161) a93161, sum(a93162) a93162, sum(a93163) a93163,  
        sum(a93171) a93171, sum(a93172) a93172, sum(a93173) a93173, sum(a93181) a93181, sum(a93182) a93182, sum(a93183) a93183,  
        sum(a93191) a93191, sum(a93192) a93192, sum(a93193) a93193, sum(a931a1) a931a1, sum(a931a2) a931a2, sum(a931a3) a931a3,  
        sum(a931b1) a931b1, sum(a931b2) a931b2, sum(a931b3) a931b3, sum(a931c1) a931c1, sum(a931c2) a931c2, sum(a931c3) a931c3,  
        sum(a9311) a9311, sum(a9312) a9312, sum(a9313) a9313, 

        sum(a93211) a93211, sum(a93212) a93212, sum(a93213) a93213, sum(a93221) a93221, sum(a93222) a93222, sum(a93223) a93223,  
        sum(a9321) a9321, sum(a9322) a9322, sum(a9323) a9323, 

        sum(a93311) a93311, sum(a93312) a93312, sum(a93313) a93313, sum(a93321) a93321, sum(a93322) a93322, sum(a93323) a93323,  
        sum(a93331) a93331, sum(a93332) a93332, sum(a93333) a93333, sum(a93341) a93341, sum(a93342) a93342, sum(a93343) a93343,  
        sum(a93351) a93351, sum(a93352) a93352, sum(a93353) a93353, sum(a93361) a93361, sum(a93362) a93362, sum(a93363) a93363,  
        sum(a93371) a93371, sum(a93372) a93372, sum(a93373) a93373, sum(a93381) a93381, sum(a93382) a93382, sum(a93383) a93383,  
        sum(a93391) a93391, sum(a93392) a93392, sum(a93393) a93393, sum(a933a1) a933a1, sum(a933a2) a933a2, sum(a933a3) a933a3, 
        sum(a9331) a9331, sum(a9332) a9332, sum(a9333) a9333,

        sum(a93411) a93411, sum(a93412) a93412, sum(a93413) a93413, sum(a93421) a93421, sum(a93422) a93422, sum(a93423) a93423,  
        sum(a93431) a93431, sum(a93432) a93432, sum(a93433) a93433, sum(a93441) a93441, sum(a93442) a93442, sum(a93443) a93443,  
        sum(a93451) a93451, sum(a93452) a93452, sum(a93453) a93453, sum(a93461) a93461, sum(a93462) a93462, sum(a93463) a93463,  
        sum(a93471) a93471, sum(a93472) a93472, sum(a93473) a93473, sum(a93481) a93481, sum(a93482) a93482, sum(a93483) a93483,  
        sum(a93491) a93491, sum(a93492) a93492, sum(a93493) a93493, sum(a934a1) a934a1, sum(a934a2) a934a2, sum(a934a3) a934a3,  
        sum(a934b1) a934b1, sum(a934b2) a934b2, sum(a934b3) a934b3, sum(a934c1) a934c1, sum(a934c2) a934c2, sum(a934c3) a934c3,  
        sum(a9341) a9341, sum(a9342) a9342, sum(a9343) a9343, 

        sum(a93511) a93511, sum(a93512) a93512, sum(a93513) a93513, sum(a93521) a93521, sum(a93522) a93522, sum(a93523) a93523,  
        sum(a93531) a93531, sum(a93532) a93532, sum(a93533) a93533, sum(a93541) a93541, sum(a93542) a93542, sum(a93543) a93543,  
        sum(a93551) a93551, sum(a93552) a93552, sum(a93553) a93553,  

        sum(a931) a931, sum(a932) a932, sum(a933) a933, 

        sum(a94111) a94111, sum(a94112) a94112, sum(a94113) a94113, sum(a94121) a94121, sum(a94122) a94122, sum(a94123) a94123,  
        sum(a94131) a94131, sum(a94132) a94132, sum(a94133) a94133,  
        sum(a9411) a9411, sum(a9412) a9412, sum(a9413) a9413, 

        sum(a94211) a94211, sum(a94212) a94212, sum(a94213) a94213, sum(a94221) a94221, sum(a94222) a94222, sum(a94223) a94223,  
        sum(a94231) a94231, sum(a94232) a94232, sum(a94233) a94233,  
        sum(a9421) a9421, sum(a9422) a9422, sum(a9423) a9423, 

        sum(a94311) a94311, sum(a94312) a94312, sum(a94313) a94313, sum(a94321) a94321, sum(a94322) a94322, sum(a94323) a94323,  
        sum(a94331) a94331, sum(a94332) a94332, sum(a94333) a94333,  
        sum(a9431) a9431, sum(a9432) a9432, sum(a9433) a9433, 
        
        sum(a94411) a94411, sum(a94412) a94412, sum(a94413) a94413, sum(a94421) a94421, sum(a94422) a94422, sum(a94423) a94423,  
        sum(a94431) a94431, sum(a94432) a94432, sum(a94433) a94433,  
        sum(a9441) a9441, sum(a9442) a9442, sum(a9443) a9443, 

        sum(a94511) a94511, sum(a94512) a94512, sum(a94513) a94513, sum(a94521) a94521, sum(a94522) a94522, sum(a94523) a94523,  
        sum(a94531) a94531, sum(a94532) a94532, sum(a94533) a94533,  
        sum(a9451) a9451, sum(a9452) a9452, sum(a9453) a9453, 

        sum(a94611) a94611, sum(a94612) a94612, sum(a94613) a94613, sum(a94621) a94621, sum(a94622) a94622, sum(a94623) a94623,  
        sum(a94631) a94631, sum(a94632) a94632, sum(a94633) a94633,  
        sum(a9461) a9461, sum(a9462) a9462, sum(a9463) a9463,

        sum(a94711) a94711, sum(a94712) a94712, sum(a94713) a94713, sum(a94721) a94721, sum(a94722) a94722, sum(a94723) a94723,  
        sum(a94731) a94731, sum(a94732) a94732, sum(a94733) a94733, sum(a94741) a94741, sum(a94742) a94742, sum(a94743) a94743,  
        sum(a94751) a94751, sum(a94752) a94752, sum(a94753) a94753, sum(a94761) a94761, sum(a94762) a94762, sum(a94763) a94763, 
        sum(a94771) a94771, sum(a94772) a94772, sum(a94773) a94773, sum(a94781) a94781, sum(a94782) a94782, sum(a94783) a94783, 
        sum(a94791) a94791, sum(a94792) a94792, sum(a94793) a94793, sum(a947a1) a947a1, sum(a947a2) a947a2, sum(a947a3) a947a3, 

        sum(a941) a941, sum(a942) a942, sum(a943) a943

        from 
          (select a.occ02, a.ima10, a.ta_oea004,  
            decode((a.ima10),'9311',a.sfb08, 0) as a93111,
            decode((a.ima10||a.ta_oea004),'93112',a.sfb08, 0) as a93112,
            decode((a.ima10||a.ta_oea004),'93113',a.sfb08, 0) as a93113,
            decode((a.ima10),'9312',a.sfb08, 0) as a93121,
            decode((a.ima10||a.ta_oea004),'93122',a.sfb08, 0) as a93122,
            decode((a.ima10||a.ta_oea004),'93123',a.sfb08, 0) as a93123,
            decode((a.ima10),'9313',a.sfb08, 0) as a93131,
            decode((a.ima10||a.ta_oea004),'93132',a.sfb08, 0) as a93132,
            decode((a.ima10||a.ta_oea004),'93133',a.sfb08, 0) as a93133,
            decode((a.ima10),'9314',a.sfb08, 0) as a93141,
            decode((a.ima10||a.ta_oea004),'93142',a.sfb08, 0) as a93142,
            decode((a.ima10||a.ta_oea004),'93143',a.sfb08, 0) as a93143,
            decode((a.ima10),'9315',a.sfb08, 0) as a93151,
            decode((a.ima10||a.ta_oea004),'93152',a.sfb08, 0) as a93152,
            decode((a.ima10||a.ta_oea004),'93153',a.sfb08, 0) as a93153,
            decode((a.ima10),'9316',a.sfb08, 0) as a93161,
            decode((a.ima10||a.ta_oea004),'93162',a.sfb08, 0) as a93162,
            decode((a.ima10||a.ta_oea004),'93163',a.sfb08, 0) as a93163,
            decode((a.ima10),'9317',a.sfb08, 0) as a93171,
            decode((a.ima10||a.ta_oea004),'93172',a.sfb08, 0) as a93172,
            decode((a.ima10||a.ta_oea004),'93173',a.sfb08, 0) as a93173,
            decode((a.ima10),'9318',a.sfb08, 0) as a93181,
            decode((a.ima10||a.ta_oea004),'93182',a.sfb08, 0) as a93182,
            decode((a.ima10||a.ta_oea004),'93183',a.sfb08, 0) as a93183,
            decode((a.ima10),'9319',a.sfb08, 0) as a93191,
            decode((a.ima10||a.ta_oea004),'93192',a.sfb08, 0) as a93192,
            decode((a.ima10||a.ta_oea004),'93193',a.sfb08, 0) as a93193,
            
            decode((a.ima10),'931A',a.sfb08, 0) as a931a1,
            decode((a.ima10||a.ta_oea004),'931A2',a.sfb08, 0) as a931a2,
            decode((a.ima10||a.ta_oea004),'931A3',a.sfb08, 0) as a931a3,
            
            decode((a.ima10),'931B',a.sfb08, 0) as a931b1,
            decode((a.ima10||a.ta_oea004),'931B2',a.sfb08, 0) as a931b2,
            decode((a.ima10||a.ta_oea004),'931B3',a.sfb08, 0) as a931b3,
            
            decode((a.ima10),'931C',a.sfb08, 0) as a931c1,
            decode((a.ima10||a.ta_oea004),'931C2',a.sfb08, 0) as a931c2,
            decode((a.ima10||a.ta_oea004),'931C3',a.sfb08, 0) as a931c3,
            
            decode((substr(a.ima10,1,3)),'931',a.sfb08, 0) as a9311,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9312',a.sfb08, 0) as a9312,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9313',a.sfb08, 0) as a9313,

            decode((a.ima10),'9321',a.sfb08, 0) as a93211,
            decode((a.ima10||a.ta_oea004),'93212',a.sfb08, 0) as a93212,
            decode((a.ima10||a.ta_oea004),'93213',a.sfb08, 0) as a93213,
            decode((a.ima10),'9322',a.sfb08, 0) as a93221,
            decode((a.ima10||a.ta_oea004),'93222',a.sfb08, 0) as a93222,
            decode((a.ima10||a.ta_oea004),'93223',a.sfb08, 0) as a93223,
            decode((substr(a.ima10,1,3)),'932',a.sfb08, 0) as a9321,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9322',a.sfb08, 0) as a9322,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9323',a.sfb08, 0) as a9323,

            decode((a.ima10),'9331',a.sfb08, 0) as a93311,
            decode((a.ima10||a.ta_oea004),'93312',a.sfb08, 0) as a93312,
            decode((a.ima10||a.ta_oea004),'93313',a.sfb08, 0) as a93313,
            decode((a.ima10),'9332',a.sfb08, 0) as a93321,
            decode((a.ima10||a.ta_oea004),'93322',a.sfb08, 0) as a93322,
            decode((a.ima10||a.ta_oea004),'93323',a.sfb08, 0) as a93323,
            decode((a.ima10),'9333',a.sfb08, 0) as a93331,
            decode((a.ima10||a.ta_oea004),'93332',a.sfb08, 0) as a93332,
            decode((a.ima10||a.ta_oea004),'93333',a.sfb08, 0) as a93333,
            decode((a.ima10),'9334',a.sfb08, 0) as a93341,
            decode((a.ima10||a.ta_oea004),'93342',a.sfb08, 0) as a93342,
            decode((a.ima10||a.ta_oea004),'93343',a.sfb08, 0) as a93343,
            decode((a.ima10),'9335',a.sfb08, 0) as a93351,
            decode((a.ima10||a.ta_oea004),'93352',a.sfb08, 0) as a93352,
            decode((a.ima10||a.ta_oea004),'93353',a.sfb08, 0) as a93353,
            decode((a.ima10),'9336',a.sfb08, 0) as a93361,
            decode((a.ima10||a.ta_oea004),'93362',a.sfb08, 0) as a93362,
            decode((a.ima10||a.ta_oea004),'93363',a.sfb08, 0) as a93363,
            decode((a.ima10),'9337',a.sfb08, 0) as a93371,
            decode((a.ima10||a.ta_oea004),'93372',a.sfb08, 0) as a93372,
            decode((a.ima10||a.ta_oea004),'93373',a.sfb08, 0) as a93373,
            decode((a.ima10),'9338',a.sfb08, 0) as a93381,
            decode((a.ima10||a.ta_oea004),'93382',a.sfb08, 0) as a93382,
            decode((a.ima10||a.ta_oea004),'93383',a.sfb08, 0) as a93383,
            decode((a.ima10),'9339',a.sfb08, 0) as a93391,
            decode((a.ima10||a.ta_oea004),'93392',a.sfb08, 0) as a93392,
            decode((a.ima10||a.ta_oea004),'93393',a.sfb08, 0) as a93393,
            decode((a.ima10),'933A',a.sfb08, 0) as a933a1,  
            decode((a.ima10||a.ta_oea004),'933A2',a.sfb08, 0) as a933a2,
            decode((a.ima10||a.ta_oea004),'933A3',a.sfb08, 0) as a933a3,
            decode((substr(a.ima10,1,3)),'933',a.sfb08, 0) as a9331,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9332',a.sfb08, 0) as a9332,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9333',a.sfb08, 0) as a9333,

            decode((a.ima10),'9341',a.sfb08, 0) as a93411,
            decode((a.ima10||a.ta_oea004),'93412',a.sfb08, 0) as a93412,
            decode((a.ima10||a.ta_oea004),'93413',a.sfb08, 0) as a93413,
            decode((a.ima10),'9342',a.sfb08, 0) as a93421,
            decode((a.ima10||a.ta_oea004),'93422',a.sfb08, 0) as a93422,
            decode((a.ima10||a.ta_oea004),'93423',a.sfb08, 0) as a93423,
            decode((a.ima10),'9343',a.sfb08, 0) as a93431,
            decode((a.ima10||a.ta_oea004),'93432',a.sfb08, 0) as a93432,
            decode((a.ima10||a.ta_oea004),'93433',a.sfb08, 0) as a93433,
            decode((a.ima10),'9344',a.sfb08, 0) as a93441,
            decode((a.ima10||a.ta_oea004),'93442',a.sfb08, 0) as a93442,
            decode((a.ima10||a.ta_oea004),'93443',a.sfb08, 0) as a93443,
            decode((a.ima10),'9345',a.sfb08, 0) as a93451,
            decode((a.ima10||a.ta_oea004),'93452',a.sfb08, 0) as a93452,
            decode((a.ima10||a.ta_oea004),'93453',a.sfb08, 0) as a93453,
            decode((a.ima10),'9346',a.sfb08, 0) as a93461,
            decode((a.ima10||a.ta_oea004),'93462',a.sfb08, 0) as a93462,
            decode((a.ima10||a.ta_oea004),'93463',a.sfb08, 0) as a93463,
            decode((a.ima10),'9347',a.sfb08, 0) as a93471,
            decode((a.ima10||a.ta_oea004),'93472',a.sfb08, 0) as a93472,
            decode((a.ima10||a.ta_oea004),'93473',a.sfb08, 0) as a93473,
            decode((a.ima10),'9348',a.sfb08, 0) as a93481,
            decode((a.ima10||a.ta_oea004),'93482',a.sfb08, 0) as a93482,
            decode((a.ima10||a.ta_oea004),'93483',a.sfb08, 0) as a93483,
            decode((a.ima10),'9349',a.sfb08, 0) as a93491,
            decode((a.ima10||a.ta_oea004),'93492',a.sfb08, 0) as a93492,
            decode((a.ima10||a.ta_oea004),'93493',a.sfb08, 0) as a93493,
            decode((a.ima10),'934A',a.sfb08, 0) as a934a1,
            decode((a.ima10||a.ta_oea004),'934A2',a.sfb08, 0) as a934a2,
            decode((a.ima10||a.ta_oea004),'934A3',a.sfb08, 0) as a934a3,
            decode((a.ima10),'934B',a.sfb08, 0) as a934b1,
            decode((a.ima10||a.ta_oea004),'934B2',a.sfb08, 0) as a934b2,
            decode((a.ima10||a.ta_oea004),'934B3',a.sfb08, 0) as a934b3,
            decode((a.ima10),'934C',a.sfb08, 0) as a934c1,
            decode((a.ima10||a.ta_oea004),'934C2',a.sfb08, 0) as a934c2,
            decode((a.ima10||a.ta_oea004),'934C3',a.sfb08, 0) as a934c3,
            decode((substr(a.ima10,1,3)),'934',a.sfb08, 0) as a9341,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9342',a.sfb08, 0) as a9342,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9343',a.sfb08, 0) as a9343,

            decode((a.ima10),'9351',a.sfb08, 0) as a93511,
            decode((a.ima10||a.ta_oea004),'93512',a.sfb08, 0) as a93512,
            decode((a.ima10||a.ta_oea004),'93513',a.sfb08, 0) as a93513,
            decode((a.ima10),'9352',a.sfb08, 0) as a93521,
            decode((a.ima10||a.ta_oea004),'93522',a.sfb08, 0) as a93522,
            decode((a.ima10||a.ta_oea004),'93523',a.sfb08, 0) as a93523,
            decode((a.ima10),'9353',a.sfb08, 0) as a93531,
            decode((a.ima10||a.ta_oea004),'93532',a.sfb08, 0) as a93532,
            decode((a.ima10||a.ta_oea004),'93533',a.sfb08, 0) as a93533,
            decode((a.ima10),'9354',a.sfb08, 0) as a93541,
            decode((a.ima10||a.ta_oea004),'93542',a.sfb08, 0) as a93542,
            decode((a.ima10||a.ta_oea004),'93543',a.sfb08, 0) as a93543,
            decode((a.ima10),'9355',a.sfb08, 0) as a93551,
            decode((a.ima10||a.ta_oea004),'93552',a.sfb08, 0) as a93552,
            decode((a.ima10||a.ta_oea004),'93553',a.sfb08, 0) as a93553, 

            decode((substr(a.ima10,1,2)),'93',a.sfb08, 0) as a931,
            decode((substr(a.ima10,1,2)||a.ta_oea004),'932',a.sfb08, 0) as a932,
            decode((substr(a.ima10,1,2)||a.ta_oea004),'933',a.sfb08, 0) as a933,

            decode(a.ima10,'9411',a.sfb08, decode(a.ima01,'24111', a.sfb08, '24113', a.sfb08, '24117', a.sfb08 ,'2411B', a.sfb08, '2411C', a.sfb08, '25111', a.sfb08, '25112', a.sfb08, 0)) as a94111,
            decode((a.ima10||a.ta_oea004),'94112',a.sfb08, decode(a.ima01||a.ta_oea004,'241112', a.sfb08, '241132', a.sfb08, '241172', a.sfb08, '2411B2', a.sfb08, '2411C2', a.sfb08, '251112', a.sfb08, '251122', a.sfb08, 0)) as a94112,
            decode((a.ima10||a.ta_oea004),'94113',a.sfb08, decode(a.ima01||a.ta_oea004,'241113', a.sfb08, '241133', a.sfb08, '241173', a.sfb08, '2411B3', a.sfb08, '2411C3', a.sfb08, '251113', a.sfb08, '251123', a.sfb08, 0)) as a94113,
            decode(a.ima10,'9412',a.sfb08, decode(a.ima01,'25121', a.sfb08, '25122', a.sfb08, 0)) as a94121,
            decode((a.ima10||a.ta_oea004),'94122',a.sfb08, decode(a.ima01||a.ta_oea004,'251212', a.sfb08, '251222', a.sfb08, 0)) as a94122,
            decode((a.ima10||a.ta_oea004),'94123',a.sfb08, decode(a.ima01||a.ta_oea004,'251213', a.sfb08, '251223', a.sfb08, 0)) as a94123,
            decode((a.ima10),'9413',a.sfb08, 0) as a94131,
            decode((a.ima10||a.ta_oea004),'94132',a.sfb08, 0) as a94132,
            decode((a.ima10||a.ta_oea004),'94133',a.sfb08, 0) as a94133,
            decode((substr(a.ima10,1,3)),'941',a.sfb08, decode(a.ima01,'24111', a.sfb08, '24113', a.sfb08, '24117', a.sfb08 ,'2411B', a.sfb08, '2411C', a.sfb08, '25111', a.sfb08, '25112', a.sfb08, '25121', a.sfb08, '25122', a.sfb08, 0)) as a9411,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9412',a.sfb08, decode(a.ima01||a.ta_oea004,'241112', a.sfb08, '241132', a.sfb08, '241172', a.sfb08 ,'2411B2', a.sfb08, '2411C2', a.sfb08, '251112', a.sfb08, '25112', a.sfb08, '251212', a.sfb08, '251222', a.sfb08, 0)) as a9412,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9413',a.sfb08, decode(a.ima01||a.ta_oea004,'241113', a.sfb08, '241133', a.sfb08, '241173', a.sfb08 ,'2411B3', a.sfb08, '2411C3', a.sfb08, '251113', a.sfb08, '25112', a.sfb08, '251213', a.sfb08, '251223', a.sfb08, 0)) as a9413,

            decode((a.ima10),'9421',a.sfb08, decode(a.ima01,'24121', a.sfb08, 0) ) as a94211,
            decode((a.ima10||a.ta_oea004),'94212',a.sfb08, decode(a.ima01||a.ta_oea004,'241212', a.sfb08,0)) as a94212,
            decode((a.ima10||a.ta_oea004),'94213',a.sfb08, decode(a.ima01||a.ta_oea004,'241213', a.sfb08,0)) as a94213,
            decode((a.ima10),'9422',a.sfb08, decode(a.ima01,'24122', a.sfb08, '24123', a.sfb08, 0)) as a94221,
            decode((a.ima10||a.ta_oea004),'94222',a.sfb08, decode(a.ima01||a.ta_oea004,'241222', a.sfb08, '241232', a.sfb08, 0)) as a94222,
            decode((a.ima10||a.ta_oea004),'94223',a.sfb08, decode(a.ima01||a.ta_oea004,'241223', a.sfb08, '241233', a.sfb08, 0)) as a94223,
            decode((a.ima10),'9423',a.sfb08, 0 ) as a94231,
            decode((a.ima10||a.ta_oea004),'94232',a.sfb08, 0) as a94232,
            decode((a.ima10||a.ta_oea004),'94233',a.sfb08, 0) as a94233,
            decode((substr(a.ima10,1,3)),'942',a.sfb08, decode(a.ima01,'24121', a.sfb08, '24122', a.sfb08, '24123', a.sfb08, 0)) as a9421,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9422',a.sfb08, decode(a.ima01||a.ta_oea004,'241212', a.sfb08, '241222', a.sfb08, '241232', a.sfb08, 0)) as a9422,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9423',a.sfb08, decode(a.ima01||a.ta_oea004,'241213', a.sfb08, '241223', a.sfb08, '241233', a.sfb08, 0)) as a9423,

            decode((a.ima10),'9431',a.sfb08, 0) as a94311,
            decode((a.ima10||a.ta_oea004),'94312',a.sfb08, 0) as a94312,
            decode((a.ima10||a.ta_oea004),'94313',a.sfb08, 0) as a94313,
            decode((a.ima10),'9432',a.sfb08, 0) as a94321,
            decode((a.ima10||a.ta_oea004),'94322',a.sfb08, 0) as a94322,
            decode((a.ima10||a.ta_oea004),'94323',a.sfb08, 0) as a94323,
            decode((a.ima10),'9433',a.sfb08, 0) as a94331,
            decode((a.ima10||a.ta_oea004),'94332',a.sfb08, 0) as a94332,
            decode((a.ima10||a.ta_oea004),'94333',a.sfb08, 0) as a94333,
            decode(substr(a.ima10,1,3),'943',a.sfb08, 0) as a9431,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9432',a.sfb08, 0) as a9432,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9433',a.sfb08, 0) as a9433,

            decode((a.ima10),'9441',a.sfb08, 0) as a94411,
            decode((a.ima10||a.ta_oea004),'94412',a.sfb08, 0) as a94412,
            decode((a.ima10||a.ta_oea004),'94413',a.sfb08, 0) as a94413,
            decode((a.ima10),'9442',a.sfb08, 0) as a94421,
            decode((a.ima10||a.ta_oea004),'94422',a.sfb08, 0) as a94422,
            decode((a.ima10||a.ta_oea004),'94423',a.sfb08, 0) as a94423,
            decode((a.ima10),'9443',a.sfb08, 0) as a94431,
            decode((a.ima10||a.ta_oea004),'94432',a.sfb08, 0) as a94432,
            decode((a.ima10||a.ta_oea004),'94433',a.sfb08, 0) as a94433,
            decode((substr(a.ima10,1,3)),'944',a.sfb08, 0) as a9441,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9442',a.sfb08, 0) as a9442,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9443',a.sfb08, 0) as a9443,

            decode((a.ima10),'9451',a.sfb08, 0) as a94511,
            decode((a.ima10||a.ta_oea004),'94512',a.sfb08, 0) as a94512,
            decode((a.ima10||a.ta_oea004),'94513',a.sfb08, 0) as a94513,
            decode((a.ima10),'9452',a.sfb08, 0) as a94521,
            decode((a.ima10||a.ta_oea004),'94522',a.sfb08, 0) as a94522,
            decode((a.ima10||a.ta_oea004),'94523',a.sfb08, 0) as a94523,
            decode((a.ima10),'9453',a.sfb08, 0) as a94531,
            decode((a.ima10||a.ta_oea004),'94532',a.sfb08, 0) as a94532,
            decode((a.ima10||a.ta_oea004),'94533',a.sfb08, 0) as a94533,
            decode((substr(a.ima10,1,3)),'945',a.sfb08, 0) as a9451,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9452',a.sfb08, 0) as a9452,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9453',a.sfb08, 0) as a9453,

            decode((a.ima10),'9461',a.sfb08, 0) as a94611,
            decode((a.ima10||a.ta_oea004),'94612',a.sfb08, 0) as a94612,
            decode((a.ima10||a.ta_oea004),'94613',a.sfb08, 0) as a94613,
            decode((a.ima10),'9462',a.sfb08, 0) as a94621,
            decode((a.ima10||a.ta_oea004),'94622',a.sfb08, 0) as a94622,
            decode((a.ima10||a.ta_oea004),'94623',a.sfb08, 0) as a94623,
            decode((a.ima10),'9463',a.sfb08, 0) as a94631,
            decode((a.ima10||a.ta_oea004),'94632',a.sfb08, 0) as a94632,
            decode((a.ima10||a.ta_oea004),'94633',a.sfb08, 0) as a94633,
            decode((substr(a.ima10,1,3)),'946',a.sfb08, 0) as a9461,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9462',a.sfb08, 0) as a9462,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9463',a.sfb08, 0) as a9463,

            decode((a.ima10),'9471',a.sfb08, 0) as a94711,
            decode((a.ima10||a.ta_oea004),'94712',a.sfb08, 0) as a94712,
            decode((a.ima10||a.ta_oea004),'94713',a.sfb08, 0) as a94713,
            decode((a.ima10),'9472',a.sfb08, 0) as a94721,
            decode((a.ima10||a.ta_oea004),'94722',a.sfb08, 0) as a94722,
            decode((a.ima10||a.ta_oea004),'94723',a.sfb08, 0) as a94723,
            decode((a.ima10),'9473',a.sfb08, 0) as a94731,
            decode((a.ima10||a.ta_oea004),'94732',a.sfb08, 0) as a94732,
            decode((a.ima10||a.ta_oea004),'94733',a.sfb08, 0) as a94733,
            decode((a.ima10),'9474',a.sfb08, 0) as a94741,
            decode((a.ima10||a.ta_oea004),'94742',a.sfb08, 0) as a94742,
            decode((a.ima10||a.ta_oea004),'94743',a.sfb08, 0) as a94743,
            decode((a.ima10),'9475',a.sfb08, 0) as a94751,
            decode((a.ima10||a.ta_oea004),'94752',a.sfb08, 0) as a94752,
            decode((a.ima10||a.ta_oea004),'94753',a.sfb08, 0) as a94753,
            decode((a.ima10),'9476',a.sfb08, 0) as a94761,
            decode((a.ima10||a.ta_oea004),'94762',a.sfb08, 0) as a94762,
            decode((a.ima10||a.ta_oea004),'94763',a.sfb08, 0) as a94763,
            decode((a.ima10),'9477',a.sfb08, 0) as a94771,
            decode((a.ima10||a.ta_oea004),'94772',a.sfb08, 0) as a94772,
            decode((a.ima10||a.ta_oea004),'94773',a.sfb08, 0) as a94773,
            decode((a.ima10),'9478',a.sfb08, 0) as a94781,
            decode((a.ima10||a.ta_oea004),'94782',a.sfb08, 0) as a94782,
            decode((a.ima10||a.ta_oea004),'94783',a.sfb08, 0) as a94783,
            decode((a.ima10),'9479',a.sfb08, 0) as a94791,
            decode((a.ima10||a.ta_oea004),'94792',a.sfb08, 0) as a94792,
            decode((a.ima10||a.ta_oea004),'94793',a.sfb08, 0) as a94793,
            decode((a.ima10),'947A',a.sfb08, 0) as a947a1,
            decode((a.ima10||a.ta_oea004),'947A2',a.sfb08, 0) as a947a2,
            decode((a.ima10||a.ta_oea004),'947A3',a.sfb08, 0) as a947a3,        

            decode(a.ima01, '24111', a.sfb08*2, '24113', a.sfb08*2, '24117', a.sfb08*2, '2411B', a.sfb08*2, '2411C', a.sfb08*2, '24121', a.sfb08*2, '24122', a.sfb08*2, '24123', a.sfb08*2, '25111', a.sfb08*2, '25112', a.sfb08*2, '25121', a.sfb08*2, '25122', a.sfb08*2 ,decode((substr(a.ima10,1,2)),'94',a.sfb08, 0)) as a941,
            decode(a.ima01||a.ta_oea004, '241112', a.sfb08*2, '241132', a.sfb08*2, '241172', a.sfb08*2, '2411B2', a.sfb08*2, '2411C2', a.sfb08*2, '241212', a.sfb08*2, '241222', a.sfb08*2, '241232', a.sfb08*2, '2511122', a.sfb08*2, '251122', a.sfb08*2, '251212', a.sfb08*2, '251222', a.sfb08*2 ,decode((substr(a.ima10,1,2)||a.ta_oea004),'942',a.sfb08, 0)) as a942,
            decode(a.ima01||a.ta_oea004, '241113', a.sfb08*2, '241133', a.sfb08*2, '241173', a.sfb08*2, '2411B3', a.sfb08*2, '2411C3', a.sfb08*2, '241213', a.sfb08*2, '241223', a.sfb08*2, '241233', a.sfb08*2, '2511123', a.sfb08*2, '251123', a.sfb08*2, '251213', a.sfb08*2, '251223', a.sfb08*2 ,decode((substr(a.ima10,1,2)||a.ta_oea004),'943',a.sfb08, 0)) as a943   
              
           from  
             ( select occ02, ima01, ima10, ta_oea004, sum(sfb08*imaud07) sfb08 from 
                 ( select occ02, ima01, ima10, imaud07, ta_oea004, sfb08 from vd110.sfb_file, vd210.oea_file, vd210.ima_file, vd210.occ_file
                  where sfb22=oea01 and oea02 between to_date('$bdate','yy/mm/dd') and  to_date('$edate','yy/mm/dd') and sfb05=ima01 $clientfilter )  
               group by occ02, ima01, ima10,ta_oea004) a)
           group by occ02 order by occ02 ";  
      //工單的日期以訂單的order date為到貨日期       
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $y=4; 
      $qtytotal=0;  
      $i=0;   
      //把要印出來的數字 依位置放到陣列中 在匯出時 直接讀取出來即可
      $t=array('    ','9311','931B','9312','9313','9314','931C','9315','9318','9319','931A','9316','9317','931','9321','9322','932',
               '9331','9332','9333','9334','9335','9336','9337','9339','933A','9338','933','9341','9342','9343','9344','9345',
               '9346','9347','9348','9349','934A','934B','934C','934','9351','9352','9353','9354','9355','93',
               '9411','9412','9413','941','9421','9422','9423','942','9431','9432','9433','943','9441','9442','9443','944',
               '9451','9452','9453','945','9461','9462','9463','946','9471','9472','9473','9474','9475','9476','9477','9478',
               '9479','947A','94');
               
      $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A3', $bdate.'--'.$edate);           
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {         
          $objPHPExcel->setActiveSheetIndex(0)
                      ->setCellValue('A'.($y), $row2['OCC02']);                
              
          for ($j=1;$j<=count($t);$j++){
              //算出X軸  第幾行
              $z= $j + 65 ;
              if ( $z > 90){ //超過90以後的 都要加一個文字
                 $a = chr(floor(($z-65)/26)+64);
                 $b = chr((($z-65) % 26)+65) ;
              } else {
                 $a="";
                 $b=chr($z); 
              }  
              
              $field='A'.$t[$j].'1';
              if ($row2[$field]==0) {
                  $v='';
              } else {
                  $v=$row2[$field];
              }
              
              //if ($field=='A9431') {
              //  $objPHPExcel->setActiveSheetIndex(0)        
              //            ->setCellValue('A300', $v);  
              //}
              $objPHPExcel->setActiveSheetIndex(0)        
                          ->setCellValue($a.$b.$y, $v); 
          }
          //最右邊的合計   
          $v= "=AU".$y."+CD".$y;          
          $objPHPExcel->setActiveSheetIndex(0)        
                      ->setCellValue('CE'.$y, $v);             
                
          $y++;
      }
      
      //最後一行的合計
      $objPHPExcel->setActiveSheetIndex(0)        
                  ->setCellValue('A'.$y,'合計'); 
      for ($j=1;$j<=count($t);$j++){  
          //算出X軸  第幾行
          $z= $j + 65 ;
          if ( $z > 90){ //超過90以後的 都要加一個文字
             $a = chr(floor(($z-65)/26)+64);
             $b = chr((($z-65) % 26)+65) ;
          } else {
             $a="";
             $b=chr($z); 
          } 
          $v="=sum(".$a.$b."4:".$a.$b.($y-1).")";
          $objPHPExcel->setActiveSheetIndex(0)        
                      ->setCellValue($a.$b.$y, $v );
      }
      //最右邊的合計    
      $v= "=AU".$y."+CD".$y;          
      $objPHPExcel->setActiveSheetIndex(0)        
                  ->setCellValue('CE'.$y, $v);             
          
   
         
      $objPHPExcel->getActiveSheet()->setTitle('MonthlyCases--by Unit');
      
      //第二個sheet放到貨組數
      $objPHPExcel->setActiveSheetIndex(1)
                  ->setCellValue('A1', ""); 
                  
     // $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      $s2="select occ02, sum(a11) a11, sum(a12) a12, sum(a13) a13, sum(a21) a21, sum(a22) a22, sum(a23) a23 from
           (select occ02,
                  decode(a.ta_oea011,'1',1,'3',1, 0) as a11,               
                  decode((a.ta_oea011||a.ta_oea004),'12',1, 0) as a12,
                  decode((a.ta_oea011||a.ta_oea004),'13',1, 0) as a13,
                  decode(a.ta_oea011,'2',1,'3',1, 0) as a21, 
                  decode((a.ta_oea011||a.ta_oea004),'22',1, 0) as a22,
                  decode((a.ta_oea011||a.ta_oea004),'23',1, 0) as a23
           from
            (select occ02, ta_oea006, ta_oea011, ta_oea004 from vd210.oea_file, vd210.occ_file where oea02 between to_date('$bdate','yy/mm/dd') 
             and to_date('$edate','yy/mm/dd') $clientfilter  )a ) group by occ02 order by occ02";   
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);         
      $y=3;   
      $t=array('    ','A1','A2');
      $objPHPExcel->setActiveSheetIndex(1)
                      ->setCellValue('A1', $bdate.'--'.$edate); 
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {    
          $objPHPExcel->setActiveSheetIndex(1)
                      ->setCellValue('A'.($y), $row2['OCC02']);                                 
        
              for ($j=1;$j<=2;$j++){                  
                  //算出X軸  第幾行
                  $z= $j + 65 ;
                  if ( $z > 90){ //超過90以後的 都要加一個文字
                     $a = chr(floor(($z-65)/26)+64);
                     $b = chr((($z-65) % 26)+65) ;
                  } else {
                     $a="";
                     $b=chr($z); 
                  }    
                  $field=$t[$j].'1';                 
                  if ($row2[$field]==0) {
                      $v='';
                  } else {
                      $v=$row2[$field];
                  }
                  $objPHPExcel->setActiveSheetIndex(1)        
                              ->setCellValue($a.$b.$y, $v); 
              }    
       
          $y++;
      }            
      $objPHPExcel->setActiveSheetIndex(1)
                  ->setCellValue('A'.$y, '合計')
                  ->setCellValue('B'.$y, '=sum(B3:B' . ($y-1) . ')')
                  ->setCellValue('C'.$y, '=sum(C3:C' . ($y-1) . ')'); 

                  
      $objPHPExcel->getActiveSheet()->setTitle('MonthlyCases--by Order');                                                                
                                                                                                                                         
      // Set active sheet index to the first sheet, so Excel opens this as the first sheet
      $objPHPExcel->setActiveSheetIndex(0);

      // Redirect output to a client’s web browser (Excel5)
      header('Content-Type: application/vnd.ms-excel');  
      header('Content-Disposition: attachment;filename="DateRangeCasesbycustomer.xls"');    
      header('Cache-Control: max-age=0'); 
      $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');  
      $objWriter->save('php://output'); 
      exit;    
}  

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內每客戶到貨合計!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee" width="60">到貨區間:</td>
          <td> 
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
            <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()">  &nbsp; &nbsp; &nbsp; &nbsp;   
            <input name="clienttype" type="radio" value="1" id="clienttype1"> <label for="clienttype1"> 收款客戶 </label>&nbsp;
            <input name="clienttype" type="radio" value="2" id="clienttype1" checked> <label for="clienttype1"> 送貨客戶 </label> &nbsp; &nbsp; &nbsp; &nbsp; 
            <input type="submit" name="submit" id="submit" value="匯出">         
          </td>
      </tr>
    </table>
  </div>
</form> 