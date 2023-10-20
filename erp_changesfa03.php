<?
  session_start();
  $pagtitle = "IT &raquo; 更改備料檔發料料號"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changesfa03.php");
  

  if ($_POST["action"] == "save") {
      $sfa03=$_POST['sfa03'];  
      $sfa27=$_POST['sfa27'];   
      $sfa12=$_POST['sfa12'];
      foreach ($_POST["sfaarray"] as $sfa01){
          if ($_POST["risok" . $sfa01] == "Y") {
              //先將要修改的資料記錄下來                                    
              $queryap   = "insert into erp_sfa03_updaterecord ( sfb01, sfb27, oldsfb03, newsfb03, username, ip ) values ( 
                          '" . $sfa01                             . "',  
                          '" . $sfa27                             . "',      
                          '" . $_POST["sfa03" . $sfa01]           . "', 
                          '" . $sfa03                             . "', 
                          '" . safetext($_SESSION['account'])     . "',       
                          '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
              $resultap= mysql_query($queryap) or die ('23 erp_sfa03_updaterecord added error. ' .mysql_error());
           
              //vd110
              $ssfa= "update sfa_file set sfa03='$sfa03', sfa12='$sfa12',sfa14='$sfa14' where sfa01='$sfa01' and sfa27='$sfa27'";
              $erp_sqlsfa=oci_parse($erp_conn1,$ssfa);
              $rs=oci_execute($erp_sqlsfa);
              if ($rs) {
                  $msg.="工單備料檔發料料號 更新成功!!";  
              } else {
                  $msg.="工單備料檔發料料號 更新失敗!!";  
              }   
          }      
      }         
      msg($msg);       
      forward("erp_changesfa03.php");                                                              
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
<p>更改工單備料檔發料料號</p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2" onsubmit="return validChangeSfb03()">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            工單單號:   
            <input name="sfa01" type="text" id="sfa01" size="20" maxlength="20"> &nbsp;&nbsp;
            BOM料號:   
            <input name="sfa27" type="text" id="sfa27" size="20" maxlength="20"> &nbsp;  &nbsp; 
            新發料料號:   
            <input name="sfa03" type="text" id="sfa03" size="20" maxlength="20"> &nbsp;  &nbsp;
            新發料單位:   
            <select name="sfa12" id="sfa12" size="1"> 
            <?
              $s1= "select gfe01 from gfe_file order by gfe01";
              $erp_sql1 = oci_parse($erp_conn,$s1 );
              oci_execute($erp_sql1);  
              while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
                  echo "<option value=" . $row1["GFE01"];  
                  if ($_GET["sfa12"] == $row1["GFE01"]) echo " selected";                  
                  echo ">" . $row1['GFE01'] . "</option>"; 
              }   
            ?>
            </select>
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
    <th>工單號碼</th>
    <th>預計開工日</th> 
    <th>產品料號</th>   
    <th>產品品名</th>   
    <th>BOM料號</th> 
    <th>BOM品名</th>
    <th>作業編號</th>
    <th>原發料料號</th>  
    <th>原發料品名</th>
    <th>新發料料號</th>  
    <th>新發料品名</th> 
    <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
  </tr>
  <?                                   
      $sfa01=$_GET['sfa01']; 
      $sfa03=$_GET['sfa03']; 
      $sfa27=$_GET['sfa27']; 
      
      //取出新的料件品名
      $ssfa03new="select ima02 from ima_file where ima01='" . $sfa03 . "' "; 
      $erp_sqlsfa03new = oci_parse($erp_conn1,$ssfa03new);
      oci_execute($erp_sqlsfa03new); 
      $rowsfa03new = oci_fetch_array($erp_sqlsfa03new, OCI_ASSOC);
      
      $ssfa="select sfa01, to_char(sfb13,'mm/dd/yy') sfb13, sfb05, ima02, sfa08, sfa27, sfa03 from sfa_file, sfb_file, ima_file 
            where sfa01='$sfa01' and sfa27='$sfa27' and sfa01=sfb01 and sfb05=ima01 "; 
      $erp_sqlsfa = oci_parse($erp_conn1,$ssfa );
      oci_execute($erp_sqlsfa); 
      while ($rowsfa = oci_fetch_array($erp_sqlsfa, OCI_ASSOC)) {    
          $ssfa27="select ima02 from ima_file where ima01='" . $rowsfa['SFA27'] . "' "; 
          $erp_sqlsfa27 = oci_parse($erp_conn1,$ssfa27 );
          oci_execute($erp_sqlsfa27); 
          $rowsfa27 = oci_fetch_array($erp_sqlsfa27, OCI_ASSOC);
          
          $ssfa03old="select ima02 from ima_file where ima01='" . $rowsfa['SFA03'] . "' "; 
          $erp_sqlsfa03old = oci_parse($erp_conn1,$ssfa03old);
          oci_execute($erp_sqlsfa03old); 
          $rowsfa03old = oci_fetch_array($erp_sqlsfa03old, OCI_ASSOC);   
      ?>    
	    <tr bgcolor="#FFFFFF"> 
		      <td><img src="i/arrow.gif" width="16" height="16">  
              <input type="hidden" name="sfaarray[]" value="<?=$rowsfa['SFA01'];?>">  </td>    
			    <td><?=$rowsfa["SFA01"];?></td>
          <td><?=$rowsfa["SFB13"];?></td>   
          <td><?=$rowsfa["SFB05"];?></td> 
          <td><?=$rowsfa["IMA02"];?></td>   
          <td><?=$rowsfa["SFA27"];?></td>  
          <td><?=$rowsfa27["IMA02"];?></td>  
          <td><?=$rowsfa["SFA08"];?></td>   
          <td><?=$rowsfa["SFA03"];?></td> 
          <td><?=$rowsfa03old["IMA02"];?></td>   
          <td><?=$sfa03;?></td> 
          <td><?=$rowsfa03new["IMA02"];?></td>   
          <td width=16>    
              <input name="risok<?=$rowsfa['SFA01'];?>" type="checkbox" id="risok" value="Y" </td>    
              <input name="sfa03<?=$rowsfa['SFA01'];?>" type="hidden" value="<?=$rowsfa["SFA03"];?>">     
          </td>  
      </tr> 
      <?  
      }   
  ?>    
  <tr>
    <td>&nbsp;</td>
    <td><input type="hidden" name="sfa27"  value="<?=$_GET['sfa27'];?>">    
        <input type="hidden" name="sfa03"  value="<?=$_GET['sfa03'];?>">     
        <input type="hidden" name="sfa12"  value="<?=$_GET['sfa12'];?>">                   
        <input type="hidden" name="action" value="save">
        <input type="submit" name="Submit" value="更新">        
    </td>
  </tr>
</table>  
</form>
