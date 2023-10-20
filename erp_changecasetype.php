<?
  session_start();
  $pagtitle = "IT &raquo; 更改海關種類"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changecasetype.php");
  
  if (is_null($_GET['oldtype'])) {
    $oldtype='9211';
  } else {
    $oldtype=$_GET['oldtype'];
  }
  if (is_null($_GET['newtype'])) {
    $newtype='9211';
  } else {
    $newtype=$_GET['newtype'];
  }
  if (is_null($_GET['bdate'])) {
    $bdate=date('Y-m-d');
  } else {
    $bdate=$_GET['bdate'];
  }
      
  if ($_POST["action"] == "save") {
      $oldtype=$_POST['oldtype'];   
      $newtype=$_POST['newtype'];           
      $msg=''; 
      foreach ($_POST["ogbarray"] as $ogb01){
          if ($_POST["risok" . $ogb01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_casetype_updaterecord ( workno, oldtype, newtype, username, ip ) values (
                          '" . safetext($ogb01)                   . "',     
                          '" . safetext($oldtype)                 . "',  
                          '" . safetext($newtype)                 . "',        
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_casetype_updaterecord added error. ' .mysql_error());
                                              
              
              //vd110
              $sogb= "update tc_ogb_file set tc_ogb007='$newtype' where tc_ogb002 = '$ogb01'";
              $erp_sqlogb=oci_parse($erp_conn1,$sogb);
              $rs=oci_execute($erp_sqlogb);
              if ($rs) {
                  $msg.="VD110 tc_ogb_file $ogb01 更新成功!!";  
              } else {
                  $msg.="VD110 tc_ogb_file訂單檔 $ogb01 更新失敗!!";  
              } 
          }      
      } 
      msg($msg);     
      forward("erp_changecasetype.php?bdate=".$bdate);                                                              
  }
  
  include("_header.php");
?>
<script language='JavaScript'>
checked = false;
function checkedAll () {
  if (checked == false) {
    checked = true
  }else{
    checked = false
  }
  
  for (var i = 0; i < document.form1.elements.length; i++) {
    var e = document.form1.elements[i];
        if (e.type == 'checkbox' && e.disabled==false) {
            e.checked = checked
        }                                                          
  }  
}
</script>
<link href="css.css" rel="stylesheet" type="text/css">
<p>更改CASE的海關秤重分類. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            出貨日期:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> &nbsp;&nbsp;  &nbsp;&nbsp;  
            舊的海關分類:   
            <select name="oldtype" id="oldtype">     
                <?
                  $s1= "select azf01,azf03 from azf_file where azf02='D' order by azf01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["AZF01"];  
                      if ($oldtype == $row1["AZF01"]) echo " selected";                  
                      echo ">" . $row1['AZF01'] ."--" .$row1["AZF03"] . "</option>"; 
                  }   
                ?>
            </select>   &nbsp;&nbsp;   &nbsp;&nbsp;  
            新的海關分類:   
            <select name="newtype" id="newtype">     
                <?
                  $s1= "select azf01,azf03 from azf_file where azf02='D' order by azf01 ";
                  $erp_sql1 = oci_parse($erp_conn1,$s1 );
                  oci_execute($erp_sql1);  
                  while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                      echo "<option value=" . $row1["AZF01"];  
                      if ($newtype == $row1["AZF01"]) echo " selected";                  
                      echo ">" . $row1['AZF01'] ."--" .$row1["AZF03"] . "</option>"; 
                  }   
                ?>
            </select>   &nbsp;&nbsp;   &nbsp;&nbsp;  
            <input type="submit" name="Submit2" value="送出" />
            </div></td>        
        </tr>
    </table>
  </div>
</form>
<? if (is_null($_GET['Submit2'])) die ; ?>
<form action="<?=$PHP_SELF;?>" method="post" name="form1">
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
  <tr>
    <th width="16">&nbsp;</th>
    <th>出貨日期</th>
    <th>秤重單號</th>        
    <th>客戶代碼</th>
    <th>工單號碼</th>  
    <th>RX #</th>      
    <th>產品代碼</th> 
    <th>顆數</th>  
    <th>舊類別</th>   
    <th>新類別</th>   
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?      
      
      $sogb="select to_char(tc_oga002,'mm-dd') outdate, tc_oga001, tc_oga004, tc_ogb002, tc_ogb011, tc_ogb005, tc_ogb006, tc_ogb007 from tc_ogb_file, tc_oga_file " . 
            "where tc_ogb007='$oldtype' and tc_oga002=to_date('$bdate','yy/mm/dd') and tc_ogb001=tc_oga001 order by tc_oga002,tc_oga001,tc_oga004            ,tc_ogb002 "; 
      $erp_sqlogb = oci_parse($erp_conn1,$sogb );
      oci_execute($erp_sqlogb); 
      while ($rowogb = oci_fetch_array($erp_sqlogb, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="ogbarray[]" value="<?=$rowogb['TC_OGB002'];?>">  </td>  
          
			    <td><?=$rowogb["OUTDATE"];?></td>
          <td><?=$rowogb["TC_OGA001"];?></td>   
          <td><?=$rowogb["TC_OGA004"];?></td> 
          <td><?=$rowogb["TC_OGB002"];?></td>  
          <td><?=$rowogb["TC_OGB011"];?></td>
          <td><?=$rowogb["TC_OGB005"];?></td>  
          <td><?=$rowogb["TC_OGB006"];?></td>  
          <td><?=$rowogb["TC_OGB007"];?></td> 
          <td><?=$newtype;?></td>    
          <td width=16><input name="risok<?=$rowogb['TC_OGB002'];?>" type="checkbox" id="risok" value="Y" </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="newtype"  value="<?=$_GET['newtype'];?>">    
        <input type="hidden" name="oldtype"  value="<?=$_GET['oldtype'];?>">                     
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
