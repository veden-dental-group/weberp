<?php
session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d');
$ydate=date('Y-m-d', strtotime("-1 days"));
$month=substr($tdate,0,7);
  
//取出今天到貨顆/組數
$s1= "select sfb82, sum(sfb08*imaud08) sfb08 " .
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
    $q1 = "update  erp_led set todayincase=$sfb08 where maker='$sfb82'";
    $r1 = mysql_query($q1);  
}  

//計算當日出貨顆/床數         
    $query = "select sum(s690000) s690000, sum(s6A1000) s6A1000, sum(s6A2000) s6A2000, sum(s6A3000) s6A3000, sum(s6A4100) s6A4100, sum(s6A4200) s6A4200, sum(s6A5000) s6A5000, sum(s6A6000) s6A6000 " .
              "from " .
              "(select if(s13='690000',round(s25*s27,2),0) s690000, if(s13='6A1000',round(s25*s27,2),0) s6A1000, if(s13='6A2000',round(s25*s27,2),0) s6A2000, if(s13='6A3000',round(s25*s27,2),0) s6A3000, " .
              "if(s13='6A5000',round(s25*s27,2),0) s6A5000, if(s13='6A6000',round(s25*s27,2),0) s6A6000, if(s13='6A4300',round(s25*s27,2),0) s6A4100, if(s13='6A4300',round(s25*s28,2),0) s6A4200 " .
              "from " .
              "(select s09, if(substr(s09,1,1)='1', s13, if(substr(s09,1,2)='27', s13, '6A4300')) s13 , s25, s26, s27, s28 from casesent where s01='$tdate' ) a ) b ";

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $q1 = "update  erp_led set todaysentcase=" . $row['s690000'] . " where maker='690000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A1000'] . " where maker='6A1000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A2000'] . " where maker='6A2000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A3000'] . " where maker='6A3000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A4100'] . " where maker='6A4100' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A4200'] . " where maker='6A4200' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A5000'] . " where maker='6A5000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set todaysentcase=" . $row['s6A6000'] . " where maker='6A6000' ";
    $r1 = mysql_query($q1);    
    
    //計算當月合計出貨顆/床數         
    $query = "select sum(s690000) s690000, sum(s6A1000) s6A1000, sum(s6A2000) s6A2000, sum(s6A3000) s6A3000, sum(s6A4100) s6A4100, sum(s6A4200) s6A4200, sum(s6A5000) s6A5000, sum(s6A6000) s6A6000 " .
              "from " .
              "(select if(s13='690000',round(s25*s27,2),0) s690000, if(s13='6A1000',round(s25*s27,2),0) s6A1000, if(s13='6A2000',round(s25*s27,2),0) s6A2000, if(s13='6A3000',round(s25*s27,2),0) s6A3000, " .
              "if(s13='6A5000',round(s25*s27,2),0) s6A5000, if(s13='6A6000',round(s25*s27,2),0) s6A6000, if(s13='6A4300',round(s25*s27,2),0) s6A4100, if(s13='6A4300',round(s25*s28,2),0) s6A4200 " .
              "from " .
              "(select s09, if(substr(s09,1,1)='1', s13, if(substr(s09,1,2)='27', s13, '6A4300')) s13 , s25, s26, s27, s28 from casesent where (date_format(s01,'%Y-%m')='$month') ) a ) b ";

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);
    $q1 = "update  erp_led set totalsentcase=" . $row['s690000'] . " where maker='690000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A1000'] . " where maker='6A1000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A2000'] . " where maker='6A2000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A3000'] . " where maker='6A3000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A4100'] . " where maker='6A4100' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A4200'] . " where maker='6A4200' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A5000'] . " where maker='6A5000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalsentcase=" . $row['s6A6000'] . " where maker='6A6000' ";
    $r1 = mysql_query($q1);         
              
     
    //計算delay數
    $query = "select sum(d690000) d690000, sum(d6A1000) d6A1000, sum(d6A2000) d6A2000, sum(d6A3000) d6A3000, sum(d6A4100) d6A4100, sum(d6A4200) d6A4200, sum(d6A5000) d6A5000, sum(d6A6000) d6A6000 " .
              "from " .
              "(select if(makercode='690000', mplus, 0) d690000, if(makercode='6A1000', mplus, 0) d6A1000, if(makercode='6A2000', mplus, 0) d6A2000, if(makercode='6A3000', mplus, 0) d6A3000, " .
              "if(makercode='6A4300', mplus, 0) d6A4100, if(makercode='6A4300', rplus, 0) d6A4200, if(makercode='6A5000', mplus, 0) d6A5000, if(makercode='6A6000', mplus, 0) d6A6000 " .
              "from " .
              "(select if(substr(dd.productcode,1,1)='1', dd.makercode,if(substr(dd.productcode,1,2)='27','6A5000','6A4300')) makercode,  round((dd.qty*dd.mplus),2) mplus,  round((dd.qty*dd.rplus),2) rplus " .
              "from delay d, delaydetail dd " .
              "where d.orderno=dd.orderno and d.tdate=dd.tdate and instr( rx,'SAMPLE')=false " .
              "and (date_format(d.tdate,'%Y-%m')='$month') and d.tdate>=d.duedate and d.status='' ".
              "and d.orderno not in (select orderno from delaydetail where productcode in ('1Z151','1Z152','1Z153')))a)b ";

    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);      
    $q1 = "update  erp_led set totaldelaycase=" . $row['d690000'] . " where maker='690000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A1000'] . " where maker='6A1000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A2000'] . " where maker='6A2000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A3000'] . " where maker='6A3000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A4100'] . " where maker='6A4100' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A4200'] . " where maker='6A4200' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A5000'] . " where maker='6A5000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totaldelaycase=" . $row['d6A6000'] . " where maker='6A6000' ";
    $r1 = mysql_query($q1); 
    
    //計算當月合計 內返/重修 顆/床數         
    $query = "select sum(rj690000) rj690000, sum(rd690000) rd690000, sum(rj6A1000) rj6A1000, sum(rd6A1000) rd6A1000, sum(rj6A2000) rj6A2000, sum(rd6A2000) rd6A2000, " .
              "sum(rj6A3000) rj6A3000, sum(rd6A3000) rd6A3000, sum(rj6A4100) rj6A4100, sum(rd6A4100) rd6A4100, sum(rj6A4200) rj6A4200, sum(rd6A4200) rd6A4200, " .
              "sum(rj6A5000) rj6A5000, sum(rd6A5000) rd6A5000, sum(rj6A6000) rj6A6000, sum(rd6A6000) rd6A6000 " .
              "from " .
              "(select if(s13='690000', reject, 0) rj690000, if(s13='690000', redo, 0) rd690000, if(s13='6A1000', reject, 0) rj6A1000, if(s13='6A1000', redo, 0) rd6A1000, " .
              "if(s13='6A2000', reject, 0) rj6A2000, if(s13='6A2000', redo, 0) rd6A2000, if(s13='6A3000', reject, 0) rj6A3000, if(s13='6A3000', redo, 0) rd6A3000, " .
              "if(s13='6A4100', reject, 0) rj6A4100, if(s13='6A4100', redo, 0) rd6A4100, if(s13='6A4200', reject, 0) rj6A4200, if(s13='6A4200', redo, 0) rd6A4200, " .
              "if(s13='6A5000', reject, 0) rj6A5000, if(s13='6A5000', redo, 0) rd6A5000, if(s13='6A6000', reject, 0) rj6A6000, if(s13='6A6000', redo, 0) rd6A6000 " .
              "from " .
              "(select rid s13, sum(rqty1) reject, sum(rqty2) redo from casereject where (date_format(rdate,'%Y-%m')='$month') group by rid ) a ) b " ;
              
    $result= mysql_query($query);            
    $row   = mysql_fetch_array($result);      
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj690000'] . " where maker='690000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A1000'] . " where maker='6A1000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A2000'] . " where maker='6A2000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A3000'] . " where maker='6A3000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A4100'] . " where maker='6A4100' ";
    $r1 = mysql_query($q1);                               
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A4200'] . " where maker='6A4200' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A5000'] . " where maker='6A5000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalrejectcase=" . $row['rj6A6000'] . " where maker='6A6000' ";
    $r1 = mysql_query($q1);          
    
    $q1 = "update  erp_led set totalredocase=" . $row['rd690000'] . " where maker='690000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A1000'] . " where maker='6A1000' ";
    $r1 = mysql_query($q1);                             
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A2000'] . " where maker='6A2000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A3000'] . " where maker='6A3000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A4100'] . " where maker='6A4100' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A4200'] . " where maker='6A4200' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A5000'] . " where maker='6A5000' ";
    $r1 = mysql_query($q1); 
    $q1 = "update  erp_led set totalredocase=" . $row['rd6A6000'] . " where maker='6A6000' ";
    $r1 = mysql_query($q1);          
    
    
    
                  

$dir = dirname(__FILE__);
require_once $dir . '/phprtflite/lib/PHPRtfLite.php';

// register PHPRtfLite class loader
PHPRtfLite::registerAutoloader();

$query="select mid,mname from maker order by mid";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result)) {
    $name=$row['mname'];
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
    $query1="select todayincase, todaysentcase, totalsentcase, totalrejectcase, totalredocase, totaldelaycase, (totalsentcase-totalrejectcase-totalredocase-totaldelaycase) moneycase from erp_led where maker='$maker' limit 1"; //取出最新的一筆
    $result1= mysql_query($query1);
    $row1 = mysql_fetch_array($result1); 
    $font1 = new PHPRtfLite_Font(18, 'Tahoma', '#00ff00'); 
    $font2 = new PHPRtfLite_Font(18, 'Tahoma', '#ff0000'); 
    $sect->writeText($name.":  ", $font1);   
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
    $sect->writeText('累計Delay:', $font1);
    $sect->writeText($row1['totaldelaycase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    $sect->writeText('有效顆數:', $font1);
    $sect->writeText($row1['moneycase'], $font2);     
    $sect->writeText('顆/床  ', $font1);
    // save rft document
    $filename="d:/led/".$maker.".rtf";
    $rtf->save($filename);
}
