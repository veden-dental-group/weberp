<?
  session_start();
  $pagtitle = "IT &raquo; 更改製處"; 
  include("_data.php");
  include("_erp.php");
  auth("erp_changedept.php");
                               
  if ($_POST["action"] == "save") {
      $gem01=$_POST['gem01']; 
      foreach ($_POST["sfbarray"] as $sfb01){
          if ($_POST["risok" . $sfb01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_dept_updaterecord ( olddept, newdept, sfbno, username, ip ) values ( 
                          '" . safetext($_POST["gem".$sfb01])     . "',  
                          '" . safetext($gem01)                   . "',      
                          '" . safetext($sfb01)                   . "', 
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('19 erp_dept_updaterecord added error. ' .mysql_error());
                                                   
              $msg='';                           
              //要由工單的sfb22, sfb221 去找訂單的單號及項次
              $ssfb="select sfb22, sfb221 from sfb_file where sfb01='$sfb01'";
              $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
              oci_execute($erp_sqlsfb); 
              $rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC);
              $sfb22=$rowsfb["SFB22"];   //訂單號
              $sfb221=$rowsfb["SFB221"]; //項次
                                                               
              //vd210
              $soea= "update oeb_file set oebud04='$gem01' where oeb01='$sfb22' and oeb03='$sfb221'";
              $erp_sqloea=oci_parse($erp_conn2,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新成功!!";  
              } else {
                  $msg.="VD210 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $soea= "update oeb_file set oebud04='$gem01' where oeb01='$sfb22' and oeb03='$sfb221'";  
              $erp_sqloea=oci_parse($erp_conn1,$soea);
              $rs=oci_execute($erp_sqloea);
              if ($rs) {
                  $msg.="VD110 oea_file訂單檔 $oea01 更新成功!!";  
              } else {
                  $msg.="VD110 oea_file訂單檔 $oea01 更新失敗!!";  
              } 
              
              //vd110
              $ssfb= "update sfb_file set sfb82='$gem01' where sfb22='$sfb22' and sfb221='$sfb221'";
              $erp_sqlsfb=oci_parse($erp_conn1,$ssfb);
              $rs=oci_execute($erp_sqlsfb);
              if ($rs) {
                  $msg.="VD110 sfb_file工單檔 $oea01 更新成功!!";  
              } else {
                  $msg.="VD110 sfb_file工單檔 $oea01 更新失敗!!";  
              }
          }      
      } 
      msg($msg);          
      forward("erp_changedept.php");                                                              
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
<p>更改訂單及工單製處. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            RX #:   
            <input name="rxno" type="text" id="rxno" size="20" maxlength="20"> &nbsp;  &nbsp;&nbsp;  &nbsp; 
            新的製處:   
            <select name="gem01" id="gem01">  
            <?
              $s1= "select gem01,gem02 from gem_file where substr(gem01,1,2) in ('69','6A') order by gem01";
              $erp_sql1 = oci_parse($erp_conn1,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["GEM01"];  
                  if ($_GET["gem01"] == $row1["GEM01"]) echo " selected";                  
                  echo ">" . $row1['GEM01'] ."--" .$row1["GEM02"] . "</option>"; 
              }   
            ?>
            </select> &nbsp;  &nbsp; &nbsp;  &nbsp; 
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
    <th>RX #</th>     
    <th>開單日期>        
    <th>工單號碼</th> 
    <th>舊製處</th>  
    <th>新製處</th> 
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?  
      $gem01=$_GET['gem01'];
      $sgem="select gem02 from gem_file where gem01='$gem01'"; 
      $erp_sqlgem = oci_parse($erp_conn1,$sgem );
      oci_execute($erp_sqlgem); 
      $rowgem = oci_fetch_array($erp_sqlgem, OCI_ASSOC);    
      
      $rxno=$_GET['rxno'];
      $ssfb="select sfbud02, to_char(sfb81,'yyyy/mm/dd') sfb81, sfb82, gem01, gem02, sfb01 from sfb_file, gem_file
            where sfbud02='$rxno' and sfb82=gem01 order by sfb01 "; 
      $erp_sqlsfb = oci_parse($erp_conn1,$ssfb );
      oci_execute($erp_sqlsfb); 
      while ($rowsfb = oci_fetch_array($erp_sqlsfb, OCI_ASSOC)) {    
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfbarray[]" value="<?=$rowsfb['SFB01'];?>"> 
              <input type="hidden" name="gem<?=$rowsfb['SFB01'];?>"" value="<?=$rowsfb['SFB82'];?>"> 
          </td>   
			    <td><?=$rowsfb["SFBUD02"];?></td>
          <td><?=$rowsfb["SFB81"];?></td>   
          <td><?=$rowsfb["SFB01"];?></td> 
          <td><?=$rowsfb["GEM01"];?>  <?=$rowsfb["GEM02"];?></td> 
          <td><?=$gem01;?> <?=$rowgem['GEM02'];?></td>           
          <td width=16><input name="risok<?=$rowsfb['SFB01'];?>" type="checkbox" id="risok" value="Y" </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="rxno"  value="<?=$_GET['rxno'];?>">    
        <input type="hidden" name="gem01"  value="<?=$_GET['gem01'];?>"> 
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
