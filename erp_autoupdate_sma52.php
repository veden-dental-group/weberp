<?

//每月一日自動更新 製造系統的期數
session_start();   
$erp_db_host1 = "topprod";
$erp_db_user1 = "vd110";
$erp_db_pass1 = "vd110";  
$erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 
  
$erp_db_host2 = "topprod";
$erp_db_user2 = "vd210";
$erp_db_pass2 = "vd210";  
$erp_conn2 = oci_connect($erp_db_user2, $erp_db_pass2, $erp_db_host2,'AL32UTF8');  

$erp_db_host3 = "topprod";
$erp_db_user3 = "vd310";
$erp_db_pass3 = "vd310";  
$erp_conn3 = oci_connect($erp_db_user3, $erp_db_pass3, $erp_db_host3,'AL32UTF8');  

$erp_db_host7 = "topprod";
$erp_db_user7 = "vd710";
$erp_db_pass7 = "vd710";  
$erp_conn7 = oci_connect($erp_db_user7, $erp_db_pass7, $erp_db_host7,'AL32UTF8');  


$erp_db_host9 = "topprod";
$erp_db_user9 = "vd910";
$erp_db_pass9 = "vd910";  
$erp_conn9 = oci_connect($erp_db_user9, $erp_db_pass9, $erp_db_host9,'AL32UTF8');  


$erp_db_hosta = "topprod";
$erp_db_usera = "vda10";
$erp_db_passa = "vda10";  
$erp_conna = oci_connect($erp_db_usera, $erp_db_passa, $erp_db_hosta,'AL32UTF8');  

$erp_db_hostb = "topprod";
$erp_db_userb = "vdb10";
$erp_db_passb = "vdb10";  
$erp_connb = oci_connect($erp_db_userb, $erp_db_passb, $erp_db_hostb,'AL32UTF8');  

$erp_db_hoste = "topprod";
$erp_db_usere = "vde10";
$erp_db_passe = "vde10";  
$erp_conne = oci_connect($erp_db_usere, $erp_db_passe, $erp_db_hoste,'AL32UTF8');  

$erp_db_hostf = "topprod";
$erp_db_userf = "vdf10";
$erp_db_passf = "vdf10";  
$erp_connf = oci_connect($erp_db_userf, $erp_db_passf, $erp_db_hostf,'AL32UTF8');  

$erp_db_hostc1 = "topprod";
$erp_db_userc1 = "vdc1";
$erp_db_passc1 = "vdc1";  
$erp_connc1 = oci_connect($erp_db_userc1, $erp_db_passc1, $erp_db_hostc1,'AL32UTF8');  

$erp_db_hostg1 = "topprod";
$erp_db_userg1 = "vdg1";
$erp_db_passg1 = "vdg1";  
$erp_conng1 = oci_connect($erp_db_userg1, $erp_db_passg1, $erp_db_hostg1,'AL32UTF8');  

$erp_db_hostc = "topprod";
$erp_db_userc = "vdc10";
$erp_db_passc = "vdc10";  
$erp_connc = oci_connect($erp_db_userc, $erp_db_passc, $erp_db_hostc,'AL32UTF8');  

$erp_db_hostg = "topprod";
$erp_db_userg = "vdg10";
$erp_db_passg = "vdg10";  
$erp_conng = oci_connect($erp_db_userg, $erp_db_passg, $erp_db_hostg,'AL32UTF8');  

$erp_db_host10 = "topprod";
$erp_db_user10 = "vd10";
$erp_db_pass10 = "vd10";  
$erp_conn10 = oci_connect($erp_db_user10, $erp_db_pass10, $erp_db_host10,'AL32UTF8');  

$erp_db_hosti = "topprod";
$erp_db_useri = "vdi10";
$erp_db_passi = "vdi10";  
$erp_conni = oci_connect($erp_db_useri, $erp_db_passi, $erp_db_hosti,'AL32UTF8');  

date_default_timezone_set('Asia/Taipei');  

$yy=date('Y');
$mm=date('m');              
          
$s2= "update sma_file set sma51=$yy, sma52=$mm where sma00=0";
$erp_sql2 = oci_parse($erp_conn1,$s2 );  
oci_execute($erp_sql2);  
                                                  
$erp_sql2 = oci_parse($erp_conn2,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conn3,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conn7,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conn9,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conna,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_connb,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conne,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_connf,$s2 );  
oci_execute($erp_sql2);         

$erp_sql2 = oci_parse($erp_connc,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conng,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_connc1,$s2 );  
oci_execute($erp_sql2);

$erp_sql2 = oci_parse($erp_conng1,$s2 );  
oci_execute($erp_sql2) ;

$erp_sql2 = oci_parse($erp_conn10,$s2 );  
oci_execute($erp_sql2) ;

$erp_sql2 = oci_parse($erp_conni,$s2 );  
oci_execute($erp_sql2) ;
?>