<?

//每天0100自動更新 vd110 工單的 sfbud13 : 應出貨日期
// 根據品代的參數 + 工單的上的根數 計算出天數 
// 日期: 到貨日期
// 同一天解傳真的CASE也要重算一次
session_start();   
//定義出每種CASE 每種種類+顆數 要幾天  
$dr=array(
          '11'=>4,  '12'=>4,  '13'=>4,  '14'=>5,  '15'=>5,  '16'=>6,  '17'=>6,  '18'=>6,  '19'=>6,  '110'=>6, '111'=>6, '112'=>6, '113'=>6, '114'=>6, '115'=>6, '116'=>6,
          '117'=>6, '118'=>6, '119'=>6, '120'=>6, '121'=>6, '122'=>6, '123'=>6, '124'=>6, '125'=>6, '126'=>6, '127'=>6, '128'=>6, '129'=>6, '130'=>6, '131'=>6, '132'=>6, 
          
          '21'=>3,  '22'=>3,  '23'=>3,  '24'=>4,  '25'=>4,  '26'=>4,  '27'=>4,  '28'=>4,  '29'=>4,  '210'=>4, '211'=>4, '212'=>4, '213'=>4, '214'=>4, '215'=>4, '216'=>4,
          '217'=>4, '218'=>4, '219'=>4, '220'=>4, '221'=>4, '222'=>4, '223'=>4, '224'=>4, '225'=>4, '226'=>4, '227'=>4, '228'=>4, '229'=>4, '230'=>4, '231'=>4, '232'=>4, 
          
          '31'=>2,  '32'=>2,  '33'=>2,  '34'=>2,  '35'=>2,  '36'=>2,  '37'=>2,  '38'=>2,  '39'=>2,  '310'=>2, '311'=>2, '312'=>2, '313'=>2, '314'=>2, '315'=>2, '316'=>2,
          '317'=>2, '318'=>2, '319'=>2, '320'=>2, '321'=>2, '322'=>2, '323'=>2, '324'=>2, '325'=>2, '326'=>2, '327'=>2, '328'=>2, '329'=>2, '330'=>2, '331'=>2, '332'=>2, 
          
          '41'=>3,  '42'=>3,  '43'=>3,  '44'=>4,  '45'=>4,  '46'=>4,  '47'=>4,  '48'=>4,  '49'=>4,  '410'=>4, '411'=>4, '412'=>4, '413'=>4, '414'=>4, '415'=>4, '416'=>4,
          '417'=>4, '418'=>4, '419'=>4, '420'=>4, '421'=>4, '422'=>4, '423'=>4, '424'=>4, '425'=>4, '426'=>4, '427'=>4, '428'=>4, '429'=>4, '430'=>4, '431'=>4, '432'=>4, 
          
          '51'=>3,  '52'=>3,  '53'=>4,  '54'=>4,  '55'=>4,  '56'=>4,  '57'=>4,  '58'=>4,  '59'=>4,  '510'=>4, '511'=>4, '512'=>4, '513'=>4, '514'=>4, '515'=>4, '516'=>4,
          '517'=>4, '518'=>4, '519'=>4, '520'=>4, '521'=>4, '522'=>4, '523'=>4, '524'=>4, '525'=>4, '526'=>4, '527'=>4, '528'=>4, '529'=>4, '530'=>4, '531'=>4, '532'=>4, 
          
          '61'=>4,  '62'=>4,  '63'=>4,  '64'=>5,  '65'=>5,  '66'=>6,  '67'=>6,  '68'=>6,  '69'=>6,  '610'=>6, '611'=>6, '612'=>6, '613'=>6, '614'=>6, '615'=>6, '616'=>6,
          '617'=>6, '618'=>6, '619'=>6, '620'=>6, '621'=>6, '622'=>6, '623'=>6, '624'=>6, '625'=>6, '626'=>6, '627'=>6, '628'=>6, '629'=>6, '630'=>6, '631'=>6, '632'=>6, 
          
          '71'=>4,  '72'=>4,  '73'=>4,  '74'=>5,  '75'=>5,  '76'=>6,  '77'=>5,  '78'=>5,  '79'=>5,  '710'=>5, '711'=>5, '712'=>5, '713'=>5, '714'=>5, '715'=>5, '716'=>5,
          '717'=>5, '718'=>5, '719'=>5, '720'=>5, '721'=>5, '722'=>5, '723'=>5, '724'=>5, '725'=>5, '726'=>5, '727'=>5, '728'=>5, '729'=>5, '730'=>5, '731'=>5, '732'=>5, 
          
          '81'=>3,  '82'=>3,  '83'=>3,  '84'=>3,  '85'=>3,  '86'=>3,  '87'=>3,  '88'=>3,  '89'=>3,  '810'=>3, '811'=>3, '812'=>3, '813'=>3, '814'=>3, '815'=>3, '816'=>3,
          '817'=>3, '818'=>3, '819'=>3, '820'=>3, '821'=>3, '822'=>3, '823'=>3, '824'=>3, '825'=>3, '826'=>3, '827'=>3, '828'=>3, '829'=>3, '830'=>3, '831'=>3, '832'=>3, 
          
          '91'=>4,  '92'=>4,  '93'=>4,  '94'=>6,  '95'=>6,  '96'=>6,  '97'=>6,  '98'=>6,  '99'=>6,  '910'=>6, '911'=>6, '912'=>6, '913'=>6, '914'=>6, '915'=>6, '916'=>6,
          '917'=>6, '918'=>6, '919'=>6, '920'=>6, '921'=>6, '922'=>6, '923'=>6, '924'=>6, '925'=>6, '926'=>6, '927'=>6, '928'=>6, '929'=>6, '930'=>6, '931'=>6, '932'=>6, 
          
          'a1'=>4,  'a2'=>4,  'a3'=>4,  'a4'=>5,  'a5'=>5,  'a6'=>6,  'a7'=>6,  'a8'=>6,  'a9'=>6,  'a10'=>6, 'a11'=>6, 'a12'=>6, 'a13'=>6, 'a14'=>6, 'a15'=>6, 'a16'=>6,
          'a17'=>6, 'a18'=>6, 'a19'=>6, 'a20'=>6, 'a21'=>6, 'a22'=>6, 'a23'=>6, 'a24'=>6, 'a25'=>6, 'a26'=>6, 'a27'=>6, 'a28'=>6, 'a29'=>6, 'a30'=>6, 'a31'=>6, 'a32'=>6, 
          
          'a1'=>4,  'a2'=>4,  'a3'=>4,  'a4'=>5,  'a5'=>5,  'a6'=>6,  'a7'=>6,  'a8'=>6,  'a9'=>6,  'a10'=>6, 'a11'=>6, 'a12'=>6, 'a13'=>6, 'a14'=>6, 'a15'=>6, 'a16'=>6,
          'a17'=>6, 'a18'=>6, 'a19'=>6, 'a20'=>6, 'a21'=>6, 'a22'=>6, 'a23'=>6, 'a24'=>6, 'a25'=>6, 'a26'=>6, 'a27'=>6, 'a28'=>6, 'a29'=>6, 'a30'=>6, 'a31'=>6, 'a32'=>6, 
          
          'a1'=>4,  'a2'=>4,  'a3'=>4,  'a4'=>5,  'a5'=>5,  'a6'=>6,  'a7'=>6,  'a8'=>6,  'a9'=>6,  'a10'=>6, 'a11'=>6, 'a12'=>6, 'a13'=>6, 'a14'=>6, 'a15'=>6, 'a16'=>6,
          'a17'=>6, 'a18'=>6, 'a19'=>6, 'a20'=>6, 'a21'=>6, 'a22'=>6, 'a23'=>6, 'a24'=>6, 'a25'=>6, 'a26'=>6, 'a27'=>6, 'a28'=>6, 'a29'=>6, 'a30'=>6, 'a31'=>6, 'a32'=>6 
          );

 $day=array('$dd');
 
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
?>