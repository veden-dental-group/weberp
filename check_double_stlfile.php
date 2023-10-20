<?php                   

//檢查有無不同日期但相同檔名的檔案 傳送email通知
session_start();        
include("_data.php");
$data='Duplicate Order:'; 
date_default_timezone_set('Asia/Taipei');  
$bdate=date('Y-m-d');
$edate=date('Y-m-d',strtotime('-3 days'));
$query="select user, edate, fname from stlfile where  fname in (select fname from stlfile where edate between '$bdate' and '$edate' group by fname having count(*)>1) order by fname,edate,user";
$result=mysql_query($query);    
while ($row=mysql_fetch_array($result)){
    $data .='<br>';
    $data .= $row['user'] . '    '.$row['edate'].'    '. $row['fname'] ; 
}

if ($data!='Duplicate Order:'){                     
         ////email
    require_once("_classes.php");
    $_SERVER['SERVER_NAME'] = 'www.vedenlabs.com';
    $regards = "Veden Dental Labs Inc.";
       
    $m = new mailer();
    $m->setMessage($regards);
    $m->setPriority( 'High' );
    $m->setFrom( "erp@vedenlabs.com", "Veden Dental Labs Inc"  );
    $m->setReplyTo( "erp@vedenlabs.com", "Veden Dental Labs Inc" );   
    //$m->setMessage($data);
    $m->setHTMLMessage($data);
    date_default_timezone_set('PRC');  
    $m->send("it@vedenlabs.com", 'Duplicate 3D Files') ;   
    $m->send("cs@vedenlabs.com", 'Duplicate 3D Files') ; 
} 


?>
