<?php
  session_start();
  $pagetitle = "財務部 &raquo; 結束配件的工單";
  include("_data.php");
  include("_erp.php");
  auth("erp_closeworkticket.php");  
  
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
  
    //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');
    
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>把配件的工單結案 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         起訖結案日期:   
         <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
         <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> 
         &nbsp;&nbsp;  
        <input type="submit" name="submit" id="submit" value="Submit">  &nbsp;&nbsp;   &nbsp;&nbsp;      
      </td></tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>訂單號</th>  
        <th>工單號</th>  
        <th>訂單開單日期</th>         
        <th>工單開單日期</th>  
        <th>產品</th>    
        <th>結案日期</th>  
        <th width="16">&nbsp;</th>   
    </tr>
    <?      
      $s2="select sfb22, to_char(sfb38,'yy/mm/dd') sfb38, to_char(sfb81,'yy/mm/dd') sfb81 from sfb_file " .
          "where sfb38 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') order by sfb22, sfb38";
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";   
      $i=1;
      $tdate=date('Y-m-d');   
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {
          $sfb22=$row2['SFB22'];
          $sfb38=$row2['SFB38'];    
          $sfb811=$row2['SFB81']; //原始的開單日期
          
          $s3="select sfb01, to_char(sfb81,'yy/mm/dd') sfb81, sfb05, ima02 from sfb_file, ima_file " . 
              "where sfb22='$sfb22' and sfb38 is null and sfb05=ima01 and (ta_ima003='Y' or ta_ima004='Y' or ta_ima005='Y') order by sfb01";
          $erp_sql3 = oci_parse($erp_conn1,$s3 );
          oci_execute($erp_sql3);  
          while ($row3 = oci_fetch_array($erp_sql3, OCI_ASSOC)) { 
              $sfb01=$row3['SFB01'];
              $sfb812=$row3['SFB81']; //本單開單日期 
              $dd='';
              $s4="update sfb_file set sfb04='8', sfb28='3', sfb38=to_date('$sfb38','yy/mm/dd'), sfbud05='closed@$tdate' where sfb01='$sfb01'"; 
              $erp_sql4 = oci_parse($erp_conn1,$s4 );
                 
              if ($sfb811!=$sfb812 && $sfb812>$sfb38 ) {  //兩個日期不一樣 且開單日>結案日表示工單有問題
                 $dd='Delete';  
                 $s5="delete from sfb_file where sfb01='$sfb01'";
                 $erp_sql5 = oci_parse($erp_conn1,$s5 );
                 oci_execute($erp_sql5);
              } else {                
                 oci_execute($erp_sql4);   
              }
          ?>                                         
              <tr bgcolor="#<?=$bgkleur;?>">    
                  <td><?=$i;?></td>  
                  <td><?=$sfb22;?></td> 
                  <td><?=$row3["SFB01"];?></td>  
                  <td><?=$sfb811;?></td>  
                  <td><?=$sfb812;?></td>    
                  <td><?=$row3['SFB05'].' -- '.$row3['IMA02'];?></td>   
                  <td><?=$sfb38;?></td> 
                  <td><?=$dd;?></td>  
              </tr>
          <? 
              $i++;
          }   
      }
    ?>     
</table>   