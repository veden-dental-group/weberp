<?
  session_start();
  $pagtitle = "IT &raquo; 更改秤重單號"; 
  include("_data.php");
  include("_erp.php");
  //auth("erp_changeweightno.php");
  

  if ($_POST["action"] == "save") {
      $newno=$_POST['newno']; 
      foreach ($_POST["tcogaarray"] as $tcoga01){
          if ($_POST["risok" . $tcoga01] == "Y") {  
              //兩個不一樣的才做更改動作
              if ($newno != $tcoga01) {
                  //先檢查是否同一天及同一個客戶才可以更新
                  $stcoga1= "select tc_oga002, tc_oga004 from tc_oga_file where tc_oga001='$newno'";
                  $erp_sqltcoga1=oci_parse($erp_conn1,$stcoga1);
                  oci_execute($erp_sqltcoga1);  
                  $rowtcoga1 = oci_fetch_array($erp_sqltcoga1, OCI_ASSOC); //取出新單的日期+客戶編號
                  
                  $stcoga2= "select tc_oga002, tc_oga004 from tc_oga_file where tc_oga001='$tcoga01'";
                  $erp_sqltcoga2=oci_parse($erp_conn1,$stcoga2);
                  oci_execute($erp_sqltcoga2);  
                  $rowtcoga2 = oci_fetch_array($erp_sqltcoga2, OCI_ASSOC); //取出新單的日期+客戶編號
                  
                  if ( ($rowtcoga1["TC_OGA002"].$rowtcoga1["TC_OGA004"]) == ($rowtcoga2["TC_OGA002"].$rowtcoga2["TC_OGA004"]) )  {
                      //更新秤重單身檔
                      $stcogb= "update tc_ogb_file set tc_ogb001='$newno' where tc_ogb001 = '$tcoga01'";
                      $erp_sqltcogb=oci_parse($erp_conn1,$stcogb);
                      oci_execute($erp_sqltcogb);
                      //刪除空的秤重單頭檔
                      $stcoga= "delete from tc_oga_file where tc_oga001 = '$tcoga01'";
                      $erp_sqltcoga=oci_parse($erp_conn1,$stcoga);
                      oci_execute($erp_sqltcoga);                   
                      $isok='Y';
                  } else {
                      msg ("秤重單號: " . $newno ."<-->". $tcoga01 . " 日期或客戶代號不同, 無法更改!!");                
                      $isok="N";
                  }  
                  
                  //先將要修改的資料記錄下來                                    
                  $queryap   = "insert into erp_weightno_updaterecord ( oldno, newno, isok, username, ip ) values ( 
                              '" . safetext($tcoga01)                 . "',  
                              '" . safetext($newno)                   . "', 
                              '" . safetext($isok)                    . "', 
                              '" . safetext($_SESSION['account'])     . "',       
                              '" . safetext($_SERVER['REMOTE_ADDR'])  . "')";
                  $resultap= mysql_query($queryap) or die ('23 erp_weightno_updaterecord added error. ' .mysql_error());   
              }
          }      
      }        
      msg ('更新完畢'); 
      forward("erp_changeweightno.php");                                                              
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
<p>更改秤重單號</p>
<form action="<?=$PHP_SELF;?>" method="post" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            新的秤重單號: 
            <input name="newno" type="text" id="newno" size="20" maxlength="20"> &nbsp;  &nbsp; 
            <input type="hidden" name="action" value="save">
            <input type="submit" name="Submit" value="更新">  
            </div></td>        
        </tr>
    </table>
  </div>           
  <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
      <th width="16">&nbsp;</th>
      <th>出貨日期</th>
      <th>客戶代號</th> 
      <th>客戶名稱</th>
      <th>秤重單號</th>  
      <th width="16"><input type="checkbox" value="checkbox" name='checkall' onclick='checkedAll();'></th>
    </tr>
    <?  
        
        $stcoga="select to_char(tc_oga002, 'mm/dd/yy') tc_oga002, tc_oga004, occ02, tc_oga001 from tc_oga_file, occ_file
                 where tc_oga004=occ01 and tc_oga002||tc_oga004 in 
                 (select tc_oga002||tc_oga004 from tc_oga_file group by tc_oga002, tc_oga004 having count(*) > 1 ) 
                 order by tc_oga002, tc_oga004";
        $erp_sqltcoga = oci_parse($erp_conn1,$stcoga );
        oci_execute($erp_sqltcoga); 
        while ($rowtcoga = oci_fetch_array($erp_sqltcoga, OCI_ASSOC)) {  
          //檢查沒有秤重的才秀出來
          $tcoga001=$rowtcoga["TC_OGA001"];
          $stcogb= "select sum(tc_ogb008) tc_ogb008 from tc_ogb_file where tc_ogb001='$tcoga001'";
          $erp_sqltcogb = oci_parse($erp_conn1,$stcogb );
          oci_execute($erp_sqltcogb); 
          $rowtcogb = oci_fetch_array($erp_sqltcogb, OCI_ASSOC);  
          if ($rowtcogb["TC_OGB008"]==0) {   
              ?>    
	            <tr bgcolor="#FFFFFF"> 
		              <td><img src="i/arrow.gif" width="16" height="16">  
                      <input type="hidden" name="tcogaarray[]" value="<?=$rowtcoga['TC_OGA001'];?>">  </td> 
			            <td><?=$rowtcoga["TC_OGA002"];?></td>
                  <td><?=$rowtcoga["TC_OGA004"];?></td>   
                  <td><?=$rowtcoga["OCC02"];?></td> 
                  <td><?=$rowtcoga["TC_OGA001"];?></td>  
                  <td width=16><input name="risok<?=$rowtcoga['TC_OGA001'];?>" type="checkbox" id="risok" value="Y" </td>  
              </tr> 
              <?  
          }        
        }   
    ?>     
  </table>  
</form>
