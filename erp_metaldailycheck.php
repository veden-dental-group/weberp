<?php
  session_start();
  $pagetitle = "業務部 &raquo; 期間內金屬重複發料檢查";
  include("_data.php");   
  auth("erp_metaldailycheck.php");  
  
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

  
  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>期間內金屬重複發料檢查 </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         扣帳日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~
        <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp; 
        &nbsp;&nbsp;      
        <input type="submit" name="submit" id="submit" value="查詢">  &nbsp;&nbsp;   &nbsp;&nbsp;              
      </td></tr>
    </table>
  </div>
</form>


<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>工單號</th>  
        <th>作業編號</th>    
        <th>料號</th>     
        <th>扣帳日期</th> 
        <th>重量</th>
           
    </tr>
    <?

      $s2= "select sfe01, sfe14, sfe07, sfe16, to_char(sfp03,'mm-dd-yyyy') sfp03 from sfe_file, sfp_file " .
           "where sfe02=sfp01 and sfe01 in (select sfe01  from sfe_file, sfp_file where sfe06='1' and sfe02=sfp01 and sfp03 between to_date('$bdate','yy/mm/dd') and to_date('$edate','yy/mm/dd') " .
           "group by sfe01, sfe14, sfe07, sfe16, sfp03 having count(*) > 1 ) ";

      $erp_sql2 = oci_parse($erp_conn,$s2 );
      oci_execute($erp_sql2);  
      $bgkleur = "ffffff";  
      $i=0;      
      $total=0;      
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {     
          $i++;                                                   
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2['SFE01'];?></td>  
                <td><?=$row2['SFE14'];?></td>
                <td><?=$row2['SFE07'];?></td> 
                <td><?=$row2['SFP03'];?></td>     
                <td style="text-align:right" ><?=number_format($row2["SFE16"],2,'.',',');?></td>   
                
            </tr>
          <?   
      }
      ?>   
</table>   