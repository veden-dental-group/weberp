<?

//重算delay的業績數
session_start();   
include("_data.php");          

$query="select pkey, productcode from delaydetail where tdate >= '2012-08-01' ";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result))  {
    $pkey=$row['pkey']  ;
    $pcode=$row['productcode'];
    $s2= "select substr(imaud02,4,3) mplus, substr(imaud02,7,3) rplus from ima_file where ima01='$pcode' ";
    $erp_sql2 = oci_parse($erp_conn1,$s2 );  
    oci_execute($erp_sql2);  
    $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
    $mplus=$row2['MPLUS'];
    $rplus=$row2['RPLUS']; 
    
    $query1="update delaydetail set mplus=$mplus, rplus=$rplus where pkey='$pkey' limit 1";
    $result1=mysql_query($query1);  
}

//重算sent的業績數 
$query="select pkey, s09 from casesent where s01>= '2012-08-01' ";
$result=mysql_query($query);
while ($row=mysql_fetch_array($result))  {
    $pkey=$row['pkey']  ;
    $pcode=$row['s09'];
    $s2= "select substr(imaud02,4,3) mplus, substr(imaud02,7,3) rplus from ima_file where ima01='$pcode' ";
    $erp_sql2 = oci_parse($erp_conn1,$s2 );  
    oci_execute($erp_sql2);  
    $row2 = oci_fetch_array($erp_sql2, OCI_ASSOC);
    $mplus=$row2['MPLUS'];
    $rplus=$row2['RPLUS']; 
    
    $query1="update casesent set s27=$mplus, s28=$rplus where pkey='$pkey' limit 1";
    $result1=mysql_query($query1);  
}    

msg('車間總出貨業績重算完畢');
forward('main.php') ;
?>