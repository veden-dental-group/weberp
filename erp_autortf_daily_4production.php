<?php
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d');
$ydate=date('Y-m-d', strtotime("-1 days"));
$month=substr($tdate,0,7);

//在table裡給每條製處增加一筆0的資料
$guid=uuid();
$query="select mid from maker order by mid";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result)) {
    $maker=$row['mid'];
    $query0="insert into erp_led (maker, guid) values ('$maker', '$maker$guid')";
    $result0=mysql_query($query0);
}


//取出今天到貨顆/組數
$s1= "select sfb82, sum(sfb08*imaud07) sfb08 " .
     "from sfb_file, ima_file, oea_file " .  
     "where sfb87='Y' " .
     "and sfb05=ima01 " .     
     "and sfb22=oea01 and oea02=to_date('$tdate','yy/mm/dd') " .
     "group by sfb82";

$erp_sql1 = oci_parse($erp_conn1,$s1 );  
oci_execute($erp_sql1);                                                     
while ($r1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
    $sfb82=$r1['SFB82'];
    $sfb08=$r1['SFB08'];
    $key  =$sfb82 . $guid;   
    $q1 = "update  erp_led set todayincase=$sfb08 where guid='$key'";
    $r1 = mysql_query($q1);  
}  

//今天寄貨, 累計寄貨, 累計重修, 累計內返
$query1= " (select s13 , sum(s25*s26) sent, 0 tsent, 0 reject, 0 redo  from casesent where s01='$tdate' group by s13) ";  
$query2= " (select s13 , 0 sent, sum(s25*s26) tsent, 0 reject, 0 redo  from casesent where (date_format(s01,'%Y-%m')='$month') group by s13) ";  
$query3= " (select rid s13, 0 sent, 0 tsent, sum(rqty1) reject, sum(rqty2) redo from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid) ";    
$query = "select s13 , sum(sent) sent, sum(tsent) tsent, sum(reject) reject, sum(redo) redo from " . 
         "(select s13 , sent, tsent, reject, redo from " .
         "( $query1 union all $query2 union all $query3 ) a) b, maker where s13=mid group by mname order by mname  ";
$result= mysql_query($query);
while ($row = mysql_fetch_array($result)){    
    $key=$row['s13'].$guid;
    $todaysent=$row['sent'];
    $totalsent=$row['tsent'];
    $totalreject=$row['reject'];
    $totalredo=$row['redo'];  
    $q1 = "update  erp_led set todaysentcase=$todaysent, totalsentcase=$totalsent, totalrejectcase=$totalreject, totalredocase=$totalredo where guid='$key'";        
    $r1 = mysql_query($q1) ;    
} 
        

$dir = dirname(__FILE__);
require_once $dir . '/phprtflite/lib/PHPRtfLite.php';

// register PHPRtfLite class loader
PHPRtfLite::registerAutoloader();

$query="select mid from maker order by mid";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result)) {
    $maker=$row['mid'];
    //Rtf document
    $rtf = new PHPRtfLite();                                 
    //Font
    $times12 = new PHPRtfLite_Font(12, 'Times new Roman');   
    //Section
    $sect = $rtf->addSection();
    //Write utf-8 encoded text.
    //Text is from file. But you can use another resouce: db, sockets and other
    //$sect->writeText(file_get_contents($dir . '/sources/utf8.txt'), $times12, null);
    //昨日到貨:xxxx  昨日出貨:xxxx  昨日delay:xxxx 累計出貨:xxxx 累計delay:xxxx  累計重修:xxxx  累計內返:xxxx
    $query1="select todayincase, todaysentcase, totalsentcase, totalrejectcase, totalredocase, (totalsentcase-totalrejectcase-totalredocase) moneycase from erp_led where maker='$maker' order by created desc limit 1"; //取出最新的一筆
    $result1= mysql_query($query1);
    $row1 = mysql_fetch_array($result1); 
    $font1 = new PHPRtfLite_Font(18, 'Arial', '#00ff00','#FFFFFF'); 
    $font2 = new PHPRtfLite_Font(18, 'Arial', '#ff0000','#FFFFFF'); 
    $sect->writeText('今日到貨:', $font1);
    $sect->writeText($row1['todayincase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('今日寄貨:', $font1);
    $sect->writeText($row1['todaysentcase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('累計寄貨:', $font1);
    $sect->writeText($row1['totalsentcase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('累計內返:', $font1);
    $sect->writeText($row1['totalrejectcase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('累計重修:', $font1);
    $sect->writeText($row1['totalredocase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('有效顆數:', $font1);
    $sect->writeText($row1['moneycase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    // save rft document
    $filename="d:/led/".$maker.".rtf";
    $rtf->save($filename);
}
