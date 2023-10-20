<?php
  session_start();
  $pagetitle = "業務部 &raquo; 到貨統計";
  include("_data.php");
  include("_erp.php");
  //auth("erp_clientdailycasereport.php");  
  
  if (is_null($_GET['bdate'])) {
    $bdate =  date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }                              
  
  if (is_null($_GET['occ01'])) {
    $occ01 =  '';
  } else {
    $occ01=$_GET['occ01'];
  }             

  $IsAjax = False;  
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>每日客戶到貨統計!! </p>     
<form action="<?=$_SERVER['PHP_SELF'];?>" method="get" name="form1">
  <div align="left">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr><td bgcolor="#eeeeee">
         到貨日期:   
        <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> 
        客戶: 
        <select name="occ01" id="occ01">  
            <option value=''>全部</option>
            <?
              $s1= "select distinct a.occ07 occ07, b.occ02 occ02 from occ_file a , occ_file b where a.occ07=b.occ01 order by a.occ07 ";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["OCC07"];  
                  if ($_GET["occ01"] == $row1["OCC07"]) echo " selected";                  
                  echo ">" . $row1['OCC07'] ."--" .$row1["OCC02"] . "</option>"; 
              }   
            ?>
        </select>                                     
        <input type="submit" name="submit" id="submit" value="查詢">     
      </td></tr>
    </table>
  </div>
</form> 


<? if (is_null($_GET['submit'])) die ; ?>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>到貨日期</th>  
        <th>RX #</th>    
        <th>工單編號</th>     
        <th>訂單編號</th>
        <th>產品編號</th>  
        <th>產品名稱</th> 
        <th>顆/床數</th> 
        <th>新/重/修</th> 
        <th>分類代號</th> 
        <th>分類名稱</th> 
        <th>客戶代號</th>   
        <th>客戶名稱</th>  
        <th>94311</th>  
        <th>94312</th>     
        <th>94313</th>     
        <th>94321</th>     
        <th>94322</th>     
        <th>94323</th>     
        <th>94331</th>     
        <th>94332</th>     
        <th>94333</th>     
        <th>9431</th>   
        <th>9432</th>  
        <th>9433</th>    
    </tr>
    <?
      $occfilter='';
      if ($_GET['occ01']!='') {
          $occfilter=" and oea04 in ( select occ01 from occ_file where occ07='" . $_GET['occ01'] . "') ";     
      }                 
      $bdate1=substr($bdate,2,2).substr($bdate,5,2).substr($bdate,8,2);  
      $s2="select oea02, ta_oea006, sfb01, oea01, sfb05, ima02, ima10, ta_oea004, azf01, azf03, sfb08, occ01, occ02, 
            decode((a.ima10),'9431',a.sfb08, 0) as a94311,
            decode((a.ima10||a.ta_oea004),'94312',a.sfb08, 0) as a94312,
            decode((a.ima10||a.ta_oea004),'94313',a.sfb08, 0) as a94313,
            decode((a.ima10),'9432',a.sfb08, 0) as a94321,
            decode((a.ima10||a.ta_oea004),'94322',a.sfb08, 0) as a94322,
            decode((a.ima10||a.ta_oea004),'94323',a.sfb08, 0) as a94323,
            decode((a.ima10),'9433',a.sfb08, 0) as a94331,
            decode((a.ima10||a.ta_oea004),'94332',a.sfb08, 0) as a94332,
            decode((a.ima10||a.ta_oea004),'94333',a.sfb08, 0) as a94333,
            decode(substr(a.ima10,1,3),'943',a.sfb08, 0) as a9431,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9432',a.sfb08, 0) as a9432,
            decode((substr(a.ima10,1,3)||a.ta_oea004),'9433',a.sfb08, 0) as a9433
            
           from  
           ( select to_char(oea02,'mm-dd-yy') oea02, ta_oea006, sfb01, oea01, sfb05, ima02, ima10, ta_oea004, azf01, azf03, sfb08, occ01, occ02 from sfb_file, oea_file, ima_file, occ_file, azf_file
             where sfb22=oea01 and oea02=to_date('$bdate1','yy/mm/dd') and sfb05=ima01 and ima10=azf01 and oea04=occ01 $occfilter 
             ) a where substr(ima10,1,3)='943' ";   
      //工單的日期以訂單的order date為到貨日期       
      $erp_sql2 = oci_parse($erp_conn1,$s2 );
      oci_execute($erp_sql2);  
      $i=0;
      $total=0;
      $bgkleur='FFFFFF';
      while ($row2 = oci_fetch_array($erp_sql2, OCI_ASSOC)) {  
          $i++;
          $total+=$row2['SFB08'];
          if ($row2['TA_OEA004']=='1') {
            $taoea='新做';
          } else if ($row2['TA_OEA004']=='2') {
            $taoea='重做'; 
          } else if ($row2['TA_OEA004']=='3') {
            $taoea='修改';
          } else {
            $taoea='錯誤';
          }
          ?>
            <tr bgcolor="#<?=$bgkleur;?>">
                <td><?=$i;?></td>      
                <td><?=$row2['OEA02'];?></td>  
                <td><?=$row2['TA_OEA006'];?></td>
                <td><?=$row2['SFB01'];?></td> 
                <td><?=$row2['OEA01'];?></td> 
                <td><?=$row2['SFB05'];?></td>   
                <td><?=$row2["IMA02"];?></td>  
                <td><?=$row2["SFB08"];?></td>  
                <td><?=$taoea;?></td>  
                <td><?=$row2["AZF01"];?></td> 
                <td><?=$row2["AZF03"];?></td>
                <td><?=$row2["OCC01"];?></td> 
                <td><?=$row2["OCC02"];?></td>  
                <td><?=$row2["A94311"];?></td>  
                <td><?=$row2["A94312"];?></td> 
                <td><?=$row2["A94313"];?></td> 
                <td><?=$row2["A94321"];?></td>  
                <td><?=$row2["A94322"];?></td> 
                <td><?=$row2["A94323"];?></td> 
                <td><?=$row2["A94331"];?></td>  
                <td><?=$row2["A94332"];?></td> 
                <td><?=$row2["A94333"];?></td> 
                <td><?=$row2["A9431"];?></td>  
                <td><?=$row2["A9432"];?></td> 
                <td><?=$row2["A9433"];?></td> 
            </tr>
          <?    
      }
      ?>   
      <tr bgcolor="#<?=$bgkleur;?>">
          <td colspan="4">&nbsp;</td> 
          <td><?=$total;?></td>  
          <td colspan="3">&nbsp;</td>      
      </tr> 
</table>   