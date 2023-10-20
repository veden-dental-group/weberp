<?php
  session_start();
  $pagetitle = "業務部 &raquo; 出貨統計V7";
  include("_data.php");
  include("_erp.php");
  date_default_timezone_set('Asia/Taipei');      
  
  $bdate = $argv[1] ? $argv[1] : date('Y-m-d');                         
  //$bdate='2014-05-01';
  
  
  $query="delete from erp_dailycount where iotype='OU' and tdate='$bdate'";
  $result=mysql_query($query) or die ('11 erp_dailycount deleted error!!');
 
  $s2="select occ07, 
    sum(a81111) a81111, sum(a81112) a81112, sum(a81113) a81113, sum(a81121) a81121, sum(a81122) a81122, sum(a81123) a81123,    
    sum(a8111) a8111, sum(a8112) a8112, sum(a8113) a8113, 
    sum(a81211) a81211, sum(a81212) a81212, sum(a81213) a81213, sum(a81221) a81221, sum(a81222) a81222, sum(a81223) a81223,    
    sum(a81231) a81231, sum(a81232) a81232, sum(a81233) a81233, sum(a81241) a81241, sum(a81242) a81242, sum(a81243) a81243,
    sum(a8121) a8121, sum(a8122) a8122, sum(a8123) a8123, 
    sum(a81211) a81311, sum(a81312) a81312, sum(a81313) a81313, sum(a81321) a81321, sum(a81322) a81322, sum(a81323) a81323,    
    sum(a81331) a81331, sum(a81332) a81332, sum(a81333) a81333, 
    sum(a8131) a8131, sum(a8132) a8132, sum(a8133) a8133, 
    sum(a81411) a81411, sum(a81412) a81412, sum(a81413) a81413, sum(a81421) a81421, sum(a81422) a81422, sum(a81423) a81423,    
    sum(a81431) a81431, sum(a81432) a81432, sum(a81433) a81433, 
    sum(a8141) a8141, sum(a8142) a8142, sum(a8143) a8143, 
    sum(a81511) a81511, sum(a81512) a81512, sum(a81513) a81513, sum(a81521) a81521, sum(a81522) a81522, sum(a81523) a81523,    
    sum(a8151) a8151, sum(a8152) a8152, sum(a8153) a8153, 
    sum(a81611) a81611, sum(a81612) a81612, sum(a81613) a81613, sum(a81621) a81621, sum(a81622) a81622, sum(a81623) a81623,    
    sum(a81631) a81631, sum(a81632) a81632, sum(a81633) a81633, sum(a81641) a81641, sum(a81642) a81642, sum(a81643) a81643,    
    sum(a8161) a8161, sum(a8162) a8162, sum(a8163) a8163,           
    sum(a81Z11) a81Z11, sum(a81Z12) a81Z12, sum(a81Z13) a81Z13, sum(a81Z21) a81Z21, sum(a81Z22) a81Z22, sum(a81Z23) a81Z23,
    sum(a81Z31) a81Z31, sum(a81Z32) a81Z32, sum(a81Z33) a81Z33, sum(a81Z41) a81Z41, sum(a81Z42) a81Z42, sum(a81Z43) a81Z43, 
    sum(a81Z51) a81Z51, sum(a81Z52) a81Z52, sum(a81Z53) a81Z53, sum(a81ZZ1) a81ZZ1, sum(a81ZZ2) a81ZZ2, sum(a81ZZ3) a81ZZ3, 
    sum(a81Z1) a81Z1, sum(a81Z2) a81Z2, sum(a81Z3) a81Z3,  
    sum(a811) a811, sum(a812) a812, sum(a813) a813,   
    
    sum(a82111) a82111, sum(a82112) a82112, sum(a82113) a82113, sum(a82121) a82121, sum(a82122) a82122, sum(a82123) a82123,    
    sum(a82131) a82131, sum(a82132) a82132, sum(a82133) a82133,
    sum(a8211) a8211, sum(a8212) a8212, sum(a8213) a8213, 
    
    sum(a82211+a92211) a82211, sum(a82212+a92212) a82212, sum(a82213+a92213) a82213, sum(a82221+a92221) a82221, sum(a82222+a92222) a82222, sum(a82223+a92223) a82223,
    sum(a8221+a92211+a92221) a8221, sum(a8222+a92212+a92222) a8222, sum(a8223+a92213+a92223) a8223, 
    
    sum(a82311+a92311) a82311, sum(a82312+a92312) a82312, sum(a82313+a92313) a82313, sum(a82321) a82321, sum(a82322) a82322, sum(a82323) a82323,    
    sum(a82331+a92331) a82331, sum(a82332+a92332) a82332, sum(a82333+a92333) a82333,
    sum(a8231+a92311+a92331) a8231, sum(a8232+a92312+a92332) a8232, sum(a8233+a92313+a92333) a8233,
    
    sum(a82Z11) a82Z11, sum(a82Z12) a82Z12, sum(a82Z13) a82Z13, sum(a82Z21+a92Z21) a82Z21, sum(a82Z22+a92Z22) a82Z22, sum(a82Z23+a92Z23) a82Z23,    
    sum(a82Z31) a82Z31, sum(a82Z32) a82Z32, sum(a82Z33) a82Z33, sum(a82ZZ1) a82ZZ1, sum(a82ZZ2) a82ZZ2, sum(a82ZZ3) a82ZZ3,
    sum(a82Z1+a92Z21) a82Z1, sum(a82Z2+a92Z22) a82Z2, sum(a82Z3+a92Z23) a82Z3,  
    
    sum(a821+a92311+a92331+a92Z21+a92211+a92221) a821, sum(a822+a92312+a92332+a92Z22+a92212+a92222) a822, sum(a823+a92313+a92333+a92Z23+a92213+a92223) a823,        
    
    sum(a83111) a83111, sum(a83112) a83112, sum(a83113) a83113, sum(a83121) a83121, sum(a83122) a83122, sum(a83123) a83123,    
    sum(a83131) a83131, sum(a83132) a83132, sum(a83133) a83133, sum(a83141) a83141, sum(a83142) a83142, sum(a83143) a83143, 
    sum(a8311) a8311, sum(a8312) a8312, sum(a8313) a8313,                                                                                                                                 
    sum(a83211) a83211, sum(a83212) a83212, sum(a83213) a83213, sum(a83221) a83221, sum(a83222) a83222, sum(a83223) a83223,    
    sum(a83231) a83231, sum(a83232) a83232, sum(a83233) a83233, sum(a83241) a83241, sum(a83242) a83242, sum(a83243) a83243, 
    sum(a8321) a8321, sum(a8322) a8322, sum(a8323) a8323,  
    sum(a83Z11) a83Z11, sum(a83Z12) a83Z12, sum(a83Z13) a83Z13,
    sum(a831) a831, sum(a832) a832, sum(a833) a833    

    from 
      (select a.occ07, a.ima10, a.ta_oea004,  
        decode((a.ima10),'8111',a.sfb08, 0) as a81111,
        decode((a.ima10||a.ta_oea004),'81112',a.sfb08, 0) as a81112,
        decode((a.ima10||a.ta_oea004),'81113',a.sfb08, 0) as a81113,
        decode((a.ima10),'8112',a.sfb08, 0) as a81121,
        decode((a.ima10||a.ta_oea004),'81122',a.sfb08, 0) as a81122,
        decode((a.ima10||a.ta_oea004),'81123',a.sfb08, 0) as a81123,  
        decode((substr(a.ima10,1,3)),'811',a.sfb08, 0) as a8111,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8112',a.sfb08, 0) as a8112,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8113',a.sfb08, 0) as a8113,
        
        decode((a.ima10),'8121',a.sfb08, 0) as a81211,
        decode((a.ima10||a.ta_oea004),'81212',a.sfb08, 0) as a81212,
        decode((a.ima10||a.ta_oea004),'81213',a.sfb08, 0) as a81213,
        decode((a.ima10),'8122',a.sfb08, 0) as a81221,
        decode((a.ima10||a.ta_oea004),'81222',a.sfb08, 0) as a81222,
        decode((a.ima10||a.ta_oea004),'81223',a.sfb08, 0) as a81223,
        decode((a.ima10),'8123',a.sfb08, 0) as a81231,
        decode((a.ima10||a.ta_oea004),'81232',a.sfb08, 0) as a81232,
        decode((a.ima10||a.ta_oea004),'81233',a.sfb08, 0) as a81233,
        decode((a.ima10),'8124',a.sfb08, 0) as a81241,
        decode((a.ima10||a.ta_oea004),'81242',a.sfb08, 0) as a81242,
        decode((a.ima10||a.ta_oea004),'81243',a.sfb08, 0) as a81243,  
        decode((substr(a.ima10,1,3)),'812',a.sfb08, 0) as a8121,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8122',a.sfb08, 0) as a8122,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8123',a.sfb08, 0) as a8123,
        
        decode((a.ima10),'8131',a.sfb08, 0) as a81311,
        decode((a.ima10||a.ta_oea004),'81312',a.sfb08, 0) as a81312,
        decode((a.ima10||a.ta_oea004),'81313',a.sfb08, 0) as a81313,
        decode((a.ima10),'8132',a.sfb08, 0) as a81321,
        decode((a.ima10||a.ta_oea004),'81322',a.sfb08, 0) as a81322,
        decode((a.ima10||a.ta_oea004),'81323',a.sfb08, 0) as a81323,
        decode((a.ima10),'8133',a.sfb08, 0) as a81331,
        decode((a.ima10||a.ta_oea004),'81332',a.sfb08, 0) as a81332,
        decode((a.ima10||a.ta_oea004),'81333',a.sfb08, 0) as a81333,      
        decode((substr(a.ima10,1,3)),'813',a.sfb08, 0) as a8131,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8132',a.sfb08, 0) as a8132,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8133',a.sfb08, 0) as a8133,
        
        decode((a.ima10),'8141',a.sfb08, 0) as a81411,
        decode((a.ima10||a.ta_oea004),'81412',a.sfb08, 0) as a81412,
        decode((a.ima10||a.ta_oea004),'81413',a.sfb08, 0) as a81413,
        decode((a.ima10),'8142',a.sfb08, 0) as a81421,
        decode((a.ima10||a.ta_oea004),'81422',a.sfb08, 0) as a81422,
        decode((a.ima10||a.ta_oea004),'81423',a.sfb08, 0) as a81423,
        decode((a.ima10),'8143',a.sfb08, 0) as a81431,
        decode((a.ima10||a.ta_oea004),'81432',a.sfb08, 0) as a81432,
        decode((a.ima10||a.ta_oea004),'81433',a.sfb08, 0) as a81433,      
        decode((substr(a.ima10,1,3)),'814',a.sfb08, 0) as a8141,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8142',a.sfb08, 0) as a8142,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8143',a.sfb08, 0) as a8143,

        decode((a.ima10),'8151',a.sfb08, 0) as a81511,
        decode((a.ima10||a.ta_oea004),'81512',a.sfb08, 0) as a81512,
        decode((a.ima10||a.ta_oea004),'81513',a.sfb08, 0) as a81513,
        decode((a.ima10),'8152',a.sfb08, 0) as a81521,
        decode((a.ima10||a.ta_oea004),'81522',a.sfb08, 0) as a81522,
        decode((a.ima10||a.ta_oea004),'81523',a.sfb08, 0) as a81523, 
        decode((substr(a.ima10,1,3)),'815',a.sfb08, 0) as a8151,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8152',a.sfb08, 0) as a8152,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8153',a.sfb08, 0) as a8153,
        
        decode((a.ima10),'8161',a.sfb08, 0) as a81611,
        decode((a.ima10||a.ta_oea004),'81612',a.sfb08, 0) as a81612,
        decode((a.ima10||a.ta_oea004),'81613',a.sfb08, 0) as a81613,
        decode((a.ima10),'8162',a.sfb08, 0) as a81621,
        decode((a.ima10||a.ta_oea004),'81622',a.sfb08, 0) as a81622,
        decode((a.ima10||a.ta_oea004),'81623',a.sfb08, 0) as a81623, 
        decode((a.ima10),'8163',a.sfb08, 0) as a81631,
        decode((a.ima10||a.ta_oea004),'81632',a.sfb08, 0) as a81632,
        decode((a.ima10||a.ta_oea004),'81633',a.sfb08, 0) as a81633, 
        decode((a.ima10),'8164',a.sfb08, 0) as a81641,
        decode((a.ima10||a.ta_oea004),'81642',a.sfb08, 0) as a81642,
        decode((a.ima10||a.ta_oea004),'81643',a.sfb08, 0) as a81643, 
        decode((substr(a.ima10,1,3)),'816',a.sfb08, 0) as a8161,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8162',a.sfb08, 0) as a8162,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8163',a.sfb08, 0) as a8163,
        
        decode((a.ima10),'81Z1',a.sfb08, 0) as a81Z11,
        decode((a.ima10||a.ta_oea004),'81Z12',a.sfb08, 0) as a81Z12,
        decode((a.ima10||a.ta_oea004),'81Z13',a.sfb08, 0) as a81Z13,           
        decode((a.ima10),'81Z2',a.sfb08, 0) as a81Z21,
        decode((a.ima10||a.ta_oea004),'81Z22',a.sfb08, 0) as a81Z22,
        decode((a.ima10||a.ta_oea004),'81Z23',a.sfb08, 0) as a81Z23, 
        decode((a.ima10),'81Z3',a.sfb08, 0) as a81Z31,
        decode((a.ima10||a.ta_oea004),'81Z32',a.sfb08, 0) as a81Z32,
        decode((a.ima10||a.ta_oea004),'81Z33',a.sfb08, 0) as a81Z33, 
        decode((a.ima10),'81Z4',a.sfb08, 0) as a81Z41,
        decode((a.ima10||a.ta_oea004),'81Z42',a.sfb08, 0) as a81Z42,
        decode((a.ima10||a.ta_oea004),'81Z43',a.sfb08, 0) as a81Z43, 
        decode((a.ima10),'81Z5',a.sfb08, 0) as a81Z51,
        decode((a.ima10||a.ta_oea004),'81Z52',a.sfb08, 0) as a81Z52,
        decode((a.ima10||a.ta_oea004),'81Z53',a.sfb08, 0) as a81Z53, 
        decode((a.ima10),'81ZZ',a.sfb08, 0) as a81ZZ1,
        decode((a.ima10||a.ta_oea004),'81ZZ2',a.sfb08, 0) as a81ZZ2,
        decode((a.ima10||a.ta_oea004),'81ZZ3',a.sfb08, 0) as a81ZZ3, 
        decode((substr(a.ima10,1,3)),'81Z',a.sfb08, 0) as a81Z1,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'81Z2',a.sfb08, 0) as a81Z2,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'81Z3',a.sfb08, 0) as a81Z3,
        
        decode((substr(a.ima10,1,2)),'81',a.sfb08, 0) as a811,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'812',a.sfb08, 0) as a812,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'813',a.sfb08, 0) as a813,    
        
        
        decode((a.ima10),'8211',a.sfb08, 0) as a82111,
        decode((a.ima10||a.ta_oea004),'82112',a.sfb08, 0) as a82112,
        decode((a.ima10||a.ta_oea004),'82113',a.sfb08, 0) as a82113,
        decode((a.ima10),'8212',a.sfb08, 0) as a82121,
        decode((a.ima10||a.ta_oea004),'82122',a.sfb08, 0) as a82122,
        decode((a.ima10||a.ta_oea004),'82123',a.sfb08, 0) as a82123,  
        decode((a.ima10),'8213',a.sfb08, 0) as a82131,
        decode((a.ima10||a.ta_oea004),'82132',a.sfb08, 0) as a82132,
        decode((a.ima10||a.ta_oea004),'82133',a.sfb08, 0) as a82133, 
        decode((substr(a.ima10,1,3)),'821',a.sfb08, 0) as a8211,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8212',a.sfb08, 0) as a8212,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8213',a.sfb08, 0) as a8213,
        
        decode((a.ima10),'8221',a.sfb08, 0) as a82211,
        decode((a.ima10||a.ta_oea004),'82212',a.sfb08, 0) as a82212,
        decode((a.ima10||a.ta_oea004),'82213',a.sfb08, 0) as a82213,
        decode((a.ima10),'8222',a.sfb08, 0) as a82221,
        decode((a.ima10||a.ta_oea004),'82222',a.sfb08, 0) as a82222,
        decode((a.ima10||a.ta_oea004),'82223',a.sfb08, 0) as a82223,  
        decode((substr(a.ima10,1,3)),'822',a.sfb08, 0) as a8221,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8222',a.sfb08, 0) as a8222,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8223',a.sfb08, 0) as a8223,
        
        decode((a.ima10),'8231',a.sfb08, 0) as a82311,
        decode((a.ima10||a.ta_oea004),'82312',a.sfb08, 0) as a82312,
        decode((a.ima10||a.ta_oea004),'82313',a.sfb08, 0) as a82313,
        decode((a.ima10),'8232',a.sfb08, 0) as a82321,
        decode((a.ima10||a.ta_oea004),'82322',a.sfb08, 0) as a82322,
        decode((a.ima10||a.ta_oea004),'82323',a.sfb08, 0) as a82323,  
        decode((a.ima10),'8233',a.sfb08, 0) as a82331,
        decode((a.ima10||a.ta_oea004),'82332',a.sfb08, 0) as a82332,
        decode((a.ima10||a.ta_oea004),'82333',a.sfb08, 0) as a82333, 
        decode((substr(a.ima10,1,3)),'823',a.sfb08, 0) as a8231,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8232',a.sfb08, 0) as a8232,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8233',a.sfb08, 0) as a8233,
        
        
        decode((a.ima01),             '24113', a.sfb08, '2411B', a.sfb08, '2411F', a.sfb08, '24123', a.sfb08, '24128', a.sfb08, '24152', a.sfb08, '25112', a.sfb08, '25122', a.sfb08, '25142', a.sfb08, '28221', a.sfb08, '28222', a.sfb08, '28223', a.sfb08, 0) as a92311,
        decode((a.ima01||a.ta_oea004),'241132',a.sfb08, '2411B2',a.sfb08, '2411F2',a.sfb08, '241232',a.sfb08, '241282',a.sfb08, '241522',a.sfb08, '241122',a.sfb08, '251222',a.sfb08, '251422',a.sfb08, '282212',a.sfb08, '282222',a.sfb08, '282232',a.sfb08, 0) as a92312,
        decode((a.ima01||a.ta_oea004),'241133',a.sfb08, '2411B3',a.sfb08, '2411F3',a.sfb08, '241233',a.sfb08, '241283',a.sfb08, '241523',a.sfb08, '241123',a.sfb08, '251223',a.sfb08, '251423',a.sfb08, '282213',a.sfb08, '282223',a.sfb08, '282233',a.sfb08, 0) as a92313,
        
        decode((a.ima01),             '24111', a.sfb08, '24117', a.sfb08, '2411C', a.sfb08, '2411E', a.sfb08, '24121', a.sfb08, '24127', a.sfb08, '24150', a.sfb08, '28211', a.sfb08, '28212', a.sfb08, '28213', a.sfb08, 0) as a92331,
        decode((a.ima01||a.ta_oea004),'241112',a.sfb08, '241172',a.sfb08, '2411C2',a.sfb08, '2411E2',a.sfb08, '241212',a.sfb08, '241272',a.sfb08, '241502',a.sfb08, '282112',a.sfb08, '282122',a.sfb08, '282132',a.sfb08, 0) as a92332,
        decode((a.ima01||a.ta_oea004),'241113',a.sfb08, '241173',a.sfb08, '2411C3',a.sfb08, '2411E3',a.sfb08, '241213',a.sfb08, '241273',a.sfb08, '241503',a.sfb08, '282113',a.sfb08, '282123',a.sfb08, '282133',a.sfb08, 0) as a92333,
        
        decode((a.ima01),             '2411A', a.sfb08, '24124', a.sfb08, '24154', a.sfb08, 0) as a92Z21,
        decode((a.ima01||a.ta_oea004),'2411A2',a.sfb08, '241242',a.sfb08, '241542',a.sfb08, 0) as a92Z22,
        decode((a.ima01||a.ta_oea004),'2411A3',a.sfb08, '241243',a.sfb08, '241543',a.sfb08, 0) as a92Z23,
        
        decode((a.ima01),             '25111', a.sfb08, '25121', a.sfb08, 0) as a92211,
        decode((a.ima01||a.ta_oea004),'251112',a.sfb08, '251212',a.sfb08, 0) as a92212,
        decode((a.ima01||a.ta_oea004),'251113',a.sfb08, '251213',a.sfb08, 0) as a92213,
        
        decode((a.ima01),             '25141', a.sfb08, 0) as a92221,
        decode((a.ima01||a.ta_oea004),'251412',a.sfb08, 0) as a92222,
        decode((a.ima01||a.ta_oea004),'251413',a.sfb08, 0) as a92223,    
       
       
        decode((a.ima10),'82Z1',a.sfb08, 0) as a82Z11,
        decode((a.ima10||a.ta_oea004),'82Z12',a.sfb08, 0) as a82Z12,
        decode((a.ima10||a.ta_oea004),'82Z13',a.sfb08, 0) as a82Z13,
        decode((a.ima10),'82Z2',a.sfb08, 0) as a82Z21,
        decode((a.ima10||a.ta_oea004),'82Z22',a.sfb08, 0) as a82Z22,
        decode((a.ima10||a.ta_oea004),'82Z23',a.sfb08, 0) as a82Z23,  
        decode((a.ima10),'82Z3',a.sfb08, 0) as a82Z31,
        decode((a.ima10||a.ta_oea004),'82Z32',a.sfb08, 0) as a82Z32,
        decode((a.ima10||a.ta_oea004),'82Z33',a.sfb08, 0) as a82Z33, 
        decode((a.ima10),'82ZZ',a.sfb08, 0) as a82ZZ1,
        decode((a.ima10||a.ta_oea004),'82ZZ2',a.sfb08, 0) as a82ZZ2,
        decode((a.ima10||a.ta_oea004),'82ZZ3',a.sfb08, 0) as a82ZZ3,
        decode((substr(a.ima10,1,3)),'82Z',a.sfb08, 0) as a82Z1,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'82Z2',a.sfb08, 0) as a82Z2,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'82Z3',a.sfb08, 0) as a82Z3,
        
        decode((substr(a.ima10,1,2)),'82',a.sfb08, 0) as a821,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'822',a.sfb08, 0) as a822,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'823',a.sfb08, 0) as a823,
        
        
        decode((a.ima10),'8311',a.sfb08, 0) as a83111,
        decode((a.ima10||a.ta_oea004),'83112',a.sfb08, 0) as a83112,
        decode((a.ima10||a.ta_oea004),'83113',a.sfb08, 0) as a83113,
        decode((a.ima10),'8312',a.sfb08, 0) as a83121,
        decode((a.ima10||a.ta_oea004),'83122',a.sfb08, 0) as a83122,
        decode((a.ima10||a.ta_oea004),'83123',a.sfb08, 0) as a83123,  
        decode((a.ima10),'8313',a.sfb08, 0) as a83131,
        decode((a.ima10||a.ta_oea004),'83132',a.sfb08, 0) as a83132,
        decode((a.ima10||a.ta_oea004),'83133',a.sfb08, 0) as a83133, 
        decode((a.ima10),'8314',a.sfb08, 0) as a83141,
        decode((a.ima10||a.ta_oea004),'83142',a.sfb08, 0) as a83142,
        decode((a.ima10||a.ta_oea004),'83143',a.sfb08, 0) as a83143, 
        decode((substr(a.ima10,1,3)),'831',a.sfb08, 0) as a8311,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8312',a.sfb08, 0) as a8312,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8313',a.sfb08, 0) as a8313,
        
        decode((a.ima10),'8321',a.sfb08, 0) as a83211,
        decode((a.ima10||a.ta_oea004),'83212',a.sfb08, 0) as a83212,
        decode((a.ima10||a.ta_oea004),'83213',a.sfb08, 0) as a83213,
        decode((a.ima10),'8322',a.sfb08, 0) as a83221,
        decode((a.ima10||a.ta_oea004),'83222',a.sfb08, 0) as a83222,
        decode((a.ima10||a.ta_oea004),'83223',a.sfb08, 0) as a83223,  
        decode((a.ima10),'8323',a.sfb08, 0) as a83231,
        decode((a.ima10||a.ta_oea004),'83232',a.sfb08, 0) as a83232,
        decode((a.ima10||a.ta_oea004),'83233',a.sfb08, 0) as a83233, 
        decode((a.ima10),'8324',a.sfb08, 0) as a83241,
        decode((a.ima10||a.ta_oea004),'83242',a.sfb08, 0) as a83242,
        decode((a.ima10||a.ta_oea004),'83243',a.sfb08, 0) as a83243, 
        decode((substr(a.ima10,1,3)),'832',a.sfb08, 0) as a8321,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8322',a.sfb08, 0) as a8322,
        decode((substr(a.ima10,1,3)||a.ta_oea004),'8323',a.sfb08, 0) as a8323,              
        
        decode((a.ima10),'83Z1',a.sfb08, 0) as a83Z11,
        decode((a.ima10||a.ta_oea004),'83Z12',a.sfb08, 0) as a83Z12,
        decode((a.ima10||a.ta_oea004),'83Z13',a.sfb08, 0) as a83Z13,
        
        decode((substr(a.ima10,1,2)),'83',a.sfb08, 0) as a831,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'832',a.sfb08, 0) as a832,
        decode((substr(a.ima10,1,2)||a.ta_oea004),'833',a.sfb08, 0) as a833            
        
       from  
         ( select occ07, ima01, ima10, ta_oea004, sum(sfb08*imaud07) sfb08 from 
             ( select occ07, ima01, imaud01 ima10, imaud07, ta_oea004, sfb08 from sfb_file, oga_file, oea_file, ima_file, occ_file 
              where sfb22=oga16 and oga16=oea01 and oga02=to_date('$bdate','yy/mm/dd') and sfb05=ima01 and oea04=occ01 )  
           group by occ07, ima01, ima10,ta_oea004) a)
       group by occ07 order by occ07 ";   
             //( select decode(occud03,'Y', occ07, occ01 ) occ07, ima01, imaud01 ima10, imaud07, ta_oea004, sfb08 from sfb_file, oga_file, oea_file, ima_file, occ_file  
  //工單的日期以訂單的order date為到貨日期       
  $erp_sql2 = oci_parse($erp_conn1,$s2 );
  oci_execute($erp_sql2);  
  $y=6; 
  $qtytotal=0;  
  $i=0;   
  //把要印出來的數字 依位置放到陣列中 在匯出時 直接讀取出來即可
  $t=array('  ','8111','8112','811','8121','8122','8123','8124','812','8131','8132','8133','813','8141','8142','8143','814','8151','8152','815','8161','8162','8163','8164','816','81Z1','81Z2','81Z3','81Z4','81Z5','81ZZ','81Z','81', 
           '8211','8212','8213','821','8221','8222','822','8231','8232','8233','823','82Z1','82Z2','82Z3','82ZZ','82Z','82',
           '8311','8312','8313','8314','831','8321','8322','8323','8324','832','83Z1','83');             
  $c=count($t) ;     
  while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
      $occ01=$row2['OCC07'];
      $socc="select occ02 from occ_file where occ01='$occ01'" ;
      $erp_sqlocc = oci_parse($erp_conn1,$socc ); 
      oci_execute($erp_sqlocc);  
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $occ02=$rowocc['OCC02'];
      $uuid=uuid();
      $query="insert into erp_dailycount ( guid, tdate, qtype, iotype, occ01, occ02 ";
      for ($j=1;$j< $c; $j++){ 
          for ($i=1;$i<4;$i++){            
              $query = $query . ',a'. $t[$j].$i;   
          }        
      }
      $query .= " ) values ( '$uuid', '$bdate', 'U', 'OU', '$occ01', '$occ02' "; 
      for ($j=1;$j< $c; $j++){ 
          for ($i=1;$i<4;$i++){      
              $value=$row2['A'.$t[$j].$i];      
              $query .= ",$value";   
          }        
      }   
      $query .= " ) ";
      $result= mysql_query($query) or die ('336 erp_dailycount added error. ' .mysql_error());
      
  }   
     
       
  //算組數 
  $s2="select occ07, sum(a11) a811, sum(a12) a812, sum(a13) a813, sum(a21) a821, sum(a22) a822, sum(a23) a823, sum(a31) a831, sum(a32) a832, sum(a33) a833 from
       (select occ07,
              decode(a.ta_oea011,'1',1,'3',1, 0) as a11,               
              decode((a.ta_oea011||a.ta_oea004),'12',1, 0) as a12,
              decode((a.ta_oea011||a.ta_oea004),'13',1, 0) as a13,
              decode(a.ta_oea011,'2',1,'3',1, 0) as a21, 
              decode((a.ta_oea011||a.ta_oea004),'22',1, 0) as a22,
              decode((a.ta_oea011||a.ta_oea004),'23',1, 0) as a23,
              decode(a.ta_oea011,'3',1,'3',1, 0) as a31, 
              decode((a.ta_oea011||a.ta_oea004),'32',1, 0) as a32,
              decode((a.ta_oea011||a.ta_oea004),'33',1, 0) as a33
       from
        (select occ07, ta_oea006, ta_oea011, ta_oea004 from oga_file, oea_file, occ_file where oga02=to_date('$bdate','yy/mm/dd') and oga04=occ01 and oga16=oea01 )a ) group by occ07 order by occ07";   
        //(select decode(occud03,'Y', occ07, occ01 ) occ07, ta_oea006, ta_oea011, ta_oea004 from oga_file, oea_file, occ_file where oga02=to_date('$bdate','yy/mm/dd') and oga04=occ01 and oga16=oea01 )a ) group by occ07 order by occ07";   
  $erp_sql2 = oci_parse($erp_conn1,$s2 );
  oci_execute($erp_sql2);         
  $y=3;   
  $t=array('    ','81','82','83');
  $c=count($t) ;
  while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
      $occ01=$row2['OCC07'];
      $socc="select occ02 from occ_file where occ01='$occ01'";
      $erp_sqlocc = oci_parse($erp_conn1,$socc ); 
      oci_execute($erp_sqlocc);  
      $rowocc = oci_fetch_array($erp_sqlocc, OCI_ASSOC);
      $occ02=$rowocc['OCC02'];
      $uuid=uuid();
      
      $query="insert into erp_dailycount ( guid, tdate, qtype, iotype, occ01, occ02 ";
      for ($j=1;$j< $c; $j++){ 
          for ($i=1;$i<4;$i++){            
              $query = $query . ',a'. $t[$j].$i;   
          }        
      }
      $query .= " ) values ( '$uuid', '$bdate', 'C', 'OU', '$occ01', '$occ02' "; 
      for ($j=1;$j< $c; $j++){ 
          for ($i=1;$i<4;$i++){      
              $value=$row2['A'.$t[$j].$i];      
              $query .= ",$value";   
          }        
      }   
      $query .= " ) ";
      $result= mysql_query($query) or die ('374 erp_dailycount added error. ' .mysql_error());
  }
  
  
  
  error_reporting(E_NONE);  
  require_once 'classes/PHPExcel.php'; 
  require_once 'classes/PHPExcel/IOFactory.php';  
  $objReader = PHPExcel_IOFactory::createReader('Excel5');
  $objPHPExcel = $objReader->load("templates/erp_dailycasesv5.xls");          
  // Add some data      
                  
  $objPHPExcel->setActiveSheetIndex(0)
              ->setCellValue('A1', substr($bdate,5,5))
              ->setCellValue('B1', "出貨統計表")
              ->setCellValue('A2', "單位：顆床數") ;  
  $objPHPExcel->setActiveSheetIndex(0)->setTitle('每日出貨顆床數'); 
  
  $query="select occ02, a811, a821, a831, a811+a821+a831 a8 from erp_dailycount where qtype='U' and iotype='OU' and tdate='$bdate' order by occ02 ";
  $result = mysql_query($query) or die ('393 erp_dailycount read error!!' . mysql_error());  
  $y=3;
  while ($row= mysql_fetch_array($result)) { 
      $objPHPExcel->setActiveSheetIndex(0)                    
                  ->setCellValue('A'.$y, $row['occ02'])   ;
      if ($row['a811']>0 ) $objPHPExcel->setActiveSheetIndex(0)           ->setCellValue('B'.$y, $row['a811']); 
      if ($row['a821']>0 ) $objPHPExcel->setActiveSheetIndex(0)           ->setCellValue('C'.$y, $row['a821']);
      if ($row['a831']>0 ) $objPHPExcel->setActiveSheetIndex(0)           ->setCellValue('D'.$y, $row['a831']);  
      $objPHPExcel->setActiveSheetIndex(0)                    //合計
                  ->setCellValue('E'.$y, '=B'.$y.'+C'.$y.'+D'.$y);  
      
      $y++;

       
  }
  $yy = $y-1;
  $objPHPExcel->setActiveSheetIndex(0)                    
              ->setCellValue('A'.$y, '合計')   
              ->setCellValue('B'.$y, '=sum(B3:B'. $yy . ')')
              ->setCellValue('C'.$y, '=sum(C3:C'. $yy . ')')  
              ->setCellValue('D'.$y, '=sum(D3:D'. $yy . ')') ; 
  $objPHPExcel->setActiveSheetIndex(0)                    //合計
              ->setCellValue('E'.$y, '=B'.$y.'+C'.$y.'+D'.$y); 

  
  $objPHPExcel->setActiveSheetIndex(1)          
              ->setCellValue('A1', "單位：顆床數") ;  
  $objPHPExcel->setActiveSheetIndex(1)->setTitle('30日內出貨顆床數'); 
  
  $tdate=$bdate;   //取出今天往前推的30天
  for ($i=1;$i <= 30; $i++){  
      $odate[$i]=$tdate;
      $tdate=date('Y-m-d',strtotime("-1 day", strtotime($tdate)));
  }
  
  $y=1;
  $objPHPExcel->setActiveSheetIndex(1)              
              ->setCellValue('C'.$y, substr($odate[1],5,5))
              ->setCellValue('D'.$y, substr($odate[2],5,5))
              ->setCellValue('E'.$y, substr($odate[3],5,5))
              ->setCellValue('F'.$y, substr($odate[4],5,5))
              ->setCellValue('G'.$y, substr($odate[5],5,5))
              ->setCellValue('H'.$y, substr($odate[6],5,5))
              ->setCellValue('I'.$y, substr($odate[7],5,5))
              ->setCellValue('J'.$y, substr($odate[8],5,5))
              ->setCellValue('K'.$y, substr($odate[9],5,5))
              ->setCellValue('L'.$y, substr($odate[10],5,5))
              ->setCellValue('M'.$y, substr($odate[11],5,5))
              ->setCellValue('N'.$y, substr($odate[12],5,5))
              ->setCellValue('O'.$y, substr($odate[13],5,5))
              ->setCellValue('P'.$y, substr($odate[14],5,5))
              ->setCellValue('Q'.$y, substr($odate[15],5,5))
              ->setCellValue('R'.$y, substr($odate[16],5,5))
              ->setCellValue('S'.$y, substr($odate[17],5,5))
              ->setCellValue('T'.$y, substr($odate[18],5,5))
              ->setCellValue('U'.$y, substr($odate[19],5,5))
              ->setCellValue('V'.$y, substr($odate[20],5,5))
              ->setCellValue('W'.$y, substr($odate[21],5,5))
              ->setCellValue('X'.$y, substr($odate[22],5,5))
              ->setCellValue('Y'.$y, substr($odate[23],5,5))
              ->setCellValue('Z'.$y, substr($odate[24],5,5))
              ->setCellValue('AA'.$y, substr($odate[25],5,5))
              ->setCellValue('AB'.$y, substr($odate[26],5,5))
              ->setCellValue('AC'.$y, substr($odate[27],5,5))
              ->setCellValue('AD'.$y, substr($odate[28],5,5))
              ->setCellValue('AE'.$y, substr($odate[29],5,5))
              ->setCellValue('AF'.$y, substr($odate[30],5,5));
  
  $tdate01=$odate[1];
  $tdate30=$odate[30];
  $query  = " select occ02 " ;
  for ($i=1; $i <=30; $i++){
      $query .= ", sum(t$i) t$i " ;   
  }
  $query .= " from ( select occ02 "; 
  for ($i=1; $i <=30; $i++){
      $query .= ", if( tdate='". $odate[$i] . "', a8,0) t" . $i ;   
  }
  $query .= " from (select occ02, tdate, (a811+a821+a831) a8 from erp_dailycount where qtype='U' and iotype='OU' and tdate <= '$tdate01' and tdate >= '$tdate30') a ) b group by occ02 order by occ02 ";
  
  $result = mysql_query($query) or die ('393 erp_dailycount read error!!' . mysql_error());  
  $y=2;
  while ($row= mysql_fetch_array($result)) { 
      $objPHPExcel->setActiveSheetIndex(1)    
                  ->setCellValue('A'.$y, $row['occ02'])
                  ->setCellValue('B'.$y, "=sum(C" . $y . ":AF" . $y . ")") ;
      if ($row['t1']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('C'.$y, $row['t1']);
      if ($row['t2']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('D'.$y, $row['t2']);                  
      if ($row['t3']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('E'.$y, $row['t3']);
      if ($row['t4']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('F'.$y, $row['t4']);
      if ($row['t5']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('G'.$y, $row['t5']);
      if ($row['t6']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('H'.$y, $row['t6']);
      if ($row['t7']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('I'.$y, $row['t7']);
      if ($row['t8']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('J'.$y, $row['t8']);
      if ($row['t9']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('K'.$y, $row['t9']);
      if ($row['t10']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('L'.$y, $row['t10']);
      if ($row['t11']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('M'.$y, $row['t11']);
      if ($row['t12']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('N'.$y, $row['t12']);
      if ($row['t13']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('O'.$y, $row['t13']);
      if ($row['t14']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('P'.$y, $row['t14']);
      if ($row['t15']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('Q'.$y, $row['t15']);
      if ($row['t16']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('R'.$y, $row['t16']);
      if ($row['t17']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('S'.$y, $row['t17']);
      if ($row['t18']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('T'.$y, $row['t18']);
      if ($row['t19']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('U'.$y, $row['t19']);
      if ($row['t20']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('V'.$y, $row['t20']);
      if ($row['t21']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('W'.$y, $row['t21']);
      if ($row['t22']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('X'.$y, $row['t22']);
      if ($row['t23']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('Y'.$y, $row['t23']);
      if ($row['t24']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('Z'.$y, $row['t24']);
      if ($row['t25']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AA'.$y, $row['t25']);
      if ($row['t26']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AB'.$y, $row['t26']);
      if ($row['t27']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AC'.$y, $row['t27']);
      if ($row['t28']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AD'.$y, $row['t28']);
      if ($row['t29']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AE'.$y, $row['t29']);
      if ($row['t30']>0 ) $objPHPExcel->setActiveSheetIndex(1)            ->setCellValue('AF'.$y, $row['t30']);   
                  
              /*    
                  ->setCellValue('D'.$y, $row['t2'])                  
                  ->setCellValue('E'.$y, $row['t3'])
                  ->setCellValue('F'.$y, $row['t4'])
                  ->setCellValue('G'.$y, $row['t5'])
                  ->setCellValue('H'.$y, $row['t6'])
                  ->setCellValue('I'.$y, $row['t7'])
                  ->setCellValue('J'.$y, $row['t8'])
                  ->setCellValue('K'.$y, $row['t9'])
                  ->setCellValue('L'.$y, $row['t10'])
                  ->setCellValue('M'.$y, $row['t11'])
                  ->setCellValue('N'.$y, $row['t12'])
                  ->setCellValue('O'.$y, $row['t13'])
                  ->setCellValue('P'.$y, $row['t14'])
                  ->setCellValue('Q'.$y, $row['t15'])
                  ->setCellValue('R'.$y, $row['t16'])
                  ->setCellValue('S'.$y, $row['t17'])
                  ->setCellValue('T'.$y, $row['t18'])
                  ->setCellValue('U'.$y, $row['t19'])
                  ->setCellValue('V'.$y, $row['t20'])
                  ->setCellValue('W'.$y, $row['t21'])
                  ->setCellValue('X'.$y, $row['t22'])
                  ->setCellValue('Y'.$y, $row['t23'])
                  ->setCellValue('Z'.$y, $row['t24'])
                  ->setCellValue('AA'.$y, $row['t25'])
                  ->setCellValue('AB'.$y, $row['t26'])
                  ->setCellValue('AC'.$y, $row['t27'])
                  ->setCellValue('AD'.$y, $row['t28'])
                  ->setCellValue('AE'.$y, $row['t29'])
                  ->setCellValue('AF'.$y, $row['t30']); 
                  
                  */
      $y++;
  }
  $yy = $y-1;
  $objPHPExcel->setActiveSheetIndex(1)                    
              ->setCellValue('A'.$y, '合計')   
              ->setCellValue('B'.$y, '=sum(B3:B'. $yy . ')')
              ->setCellValue('C'.$y, '=sum(C3:C'. $yy . ')')  
              ->setCellValue('D'.$y, '=sum(D3:D'. $yy . ')')
              ->setCellValue('E'.$y, '=sum(E3:E'. $yy . ')') 
              ->setCellValue('F'.$y, '=sum(F3:F'. $yy . ')') 
              ->setCellValue('G'.$y, '=sum(G3:H'. $yy . ')') 
              ->setCellValue('H'.$y, '=sum(H3:H'. $yy . ')') 
              ->setCellValue('I'.$y, '=sum(I3:I'. $yy . ')') 
              ->setCellValue('J'.$y, '=sum(J3:J'. $yy . ')') 
              ->setCellValue('K'.$y, '=sum(K3:K'. $yy . ')') 
              ->setCellValue('L'.$y, '=sum(L3:L'. $yy . ')') 
              ->setCellValue('M'.$y, '=sum(M3:M'. $yy . ')') 
              ->setCellValue('N'.$y, '=sum(N3:N'. $yy . ')') 
              ->setCellValue('O'.$y, '=sum(O3:O'. $yy . ')') 
              ->setCellValue('P'.$y, '=sum(P3:P'. $yy . ')') 
              ->setCellValue('Q'.$y, '=sum(Q3:Q'. $yy . ')') 
              ->setCellValue('R'.$y, '=sum(R3:R'. $yy . ')') 
              ->setCellValue('S'.$y, '=sum(S3:S'. $yy . ')') 
              ->setCellValue('T'.$y, '=sum(T3:T'. $yy . ')') 
              ->setCellValue('U'.$y, '=sum(U3:U'. $yy . ')') 
              ->setCellValue('V'.$y, '=sum(V3:V'. $yy . ')') 
              ->setCellValue('W'.$y, '=sum(W3:W'. $yy . ')') 
              ->setCellValue('X'.$y, '=sum(X3:X'. $yy . ')') 
              ->setCellValue('Y'.$y, '=sum(Y3:Y'. $yy . ')') 
              ->setCellValue('Z'.$y, '=sum(Z3:Z'. $yy . ')') 
              ->setCellValue('AA'.$y, '=sum(AA3:AA'. $yy . ')') 
              ->setCellValue('AB'.$y, '=sum(AB3:AB'. $yy . ')') 
              ->setCellValue('AC'.$y, '=sum(AC3:AC'. $yy . ')') 
              ->setCellValue('AD'.$y, '=sum(AD3:AD'. $yy . ')') 
              ->setCellValue('AE'.$y, '=sum(AE3:AE'. $yy . ')') 
              ->setCellValue('AF'.$y, '=sum(AF3:AF'. $yy . ')'); 
                                                                    
  
 
$objPHPExcel->setActiveSheetIndex(0);
    
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');   
$filename='email/' . $bdate . '_VD_DailyCasesDelivered.xls';    
$objWriter->save($filename);   

?>  
        
                            