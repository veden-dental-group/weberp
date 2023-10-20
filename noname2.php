<?php

session_start();   
include("_data.php");  
date_default_timezone_set('Asia/Taipei'); 
$tdate=date('Y-m-d');                       
$edate=date('Y-m-d', strtotime("-4 days"));   

$queryd="select orderno, orderdate from delay where tdate='$tdate' ";
$resultd=mysql_query($queryd);
while ($rowd=mysql_fetch_array($resultd)){
    $orderno=$rowd['orderno'];
    
    //取本工單各品代的合計天數
    $querydd1="select sum(makeday) makeday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd1=mysql_query($querydd1);
    $rowdd1=mysql_fetch_array($resultdd1);
    $makeday=min($rowdd1['makeday'],10);
    
    //取本工單各幾個品代
    $querydd2="select count(*) recday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd2=mysql_query($querydd2);
    $rowdd2=mysql_fetch_array($resultdd2);
    $recday=$rowdd2['recday']-1;
    
    //取本工單各品代的傳真的最大天數
    $querydd3="select max(faxday) faxday from delaydetail where tdate='$tdate' and orderno='$orderno' ";
    $resultdd3=mysql_query($querydd3);
    $rowdd3=mysql_fetch_array($resultdd3);
    $faxday=$rowdd3['faxday'];
  
    $orderdate=$rowd['orderdate'];
    $totalseconds= (min($makeday - $recday,10) + $faxday) * 86400 ;
    $duedate=date('Y-m-d',strtotime($orderdate)+$totalseconds);   
    
    $querydu="update delay set makeday=$makeday, recday=$recday, faxday=$faxday, duedate='$duedate' where tdate='$tdate' and orderno='$orderno'";
    $resultdu=mysql_query($querydu);   
}

?>
