<?php
  session_start();
  $pagetitle = "財務部 &raquo; 更改出貨單匯率";
  include("_data.php");
  auth("erp_updateexchangerate.php");  
  
  if (is_null($_GET['thismonth'])) {
    $thismonth = date("Ym",strtotime("-1 month")) ;     
  } else {      
    $thismonth = $_GET['thismonth'] ;   
  }           
  
  if ($_GET['submit']=='Update'){           
    
      $svmi="update rvb_file set rvb89='N' where rvb89 is null or rvb89='Y'";
      $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
      oci_execute($erp_sqlvmi); 
      
      $svmi="update rvv_file set rvv89='N' where rvv89 is null or rvv89='Y'";
      $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
      oci_execute($erp_sqlvmi);
      
      $svmi="update pmn_file set pmn89='N' where pmn89 is null or pmn89='Y'";
      $erp_sqlvmi = oci_parse($erp_conn1,$svmi );
      oci_execute($erp_sqlvmi);
      
      
      msg('匯率更新完畢!!');
      forward('erp_updateexchangerate.php');
      
  }
           
  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>更改出貨單匯率!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
        月份:   
        <input name="thismonth" type="text" id="thismonth" onfocus="WdatePicker({dateFmt:'yyyyMM'})" value="<?=$thismonth;?>"> &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="Update">  &nbsp;&nbsp;     
      </td></tr>
    </table>
  </div>
</form>  
