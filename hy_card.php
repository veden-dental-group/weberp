<?
  session_start();
  $pagtitle = "IT &raquo; 加班費試算"; 
  include("_data.php");
 // include("_erp.php");
 // auth("erp_changedept.php");
                               
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
  
  if (is_null($_GET['salary'])) {
    $salary = 1380;
  } else {
    $salary = $_GET['salary'];
  }                                               
  
  //for xajax
  require ('xajax/xajax_core/xajax.inc.php');
  $xajax = new xajax();
  $xajax->configure('javascript URI', 'xajax/');
  
       
    Function showLevel2($v){      
      $o=new xajaxResponse();                          
      $queryc="select distinct dept2 from hy_dept where dept1='$v' order by dept2";      
      $resultc=mysql_query($queryc);  
      if (!$resultc) {
             $o->alert("69 hy_dept error!!");   
      } else {    
          $level2 = '<select name="level2no" id="level2no" onchange="xajax_showLevel3(document.getElementById(\'level2no\').value)">';    
          $level2 .= "<option value=''>請選</option>"  ;   
          while ($rowc=mysql_fetch_array($resultc)){
            $level2 .= "<option value=" . $rowc["dept2"]  . ">" . $rowc["dept2"]."</option>";       
          }                                                                                     
          $o->assign('level2' ,"innerHTML",  $level2);                       
      }
                                                        
      $o->assign('level3' ,"innerHTML", '');     
      $o->assign('level4' ,"innerHTML", '');     
      $o->assign('level5' ,"innerHTML", ''); 
      $o->assign('level6' ,"innerHTML", ''); 
      
      return $o;
  }  
  
  Function showLevel3($v){      
      $o=new xajaxResponse();                          
      $queryc="select distinct dept3 from hy_dept where dept2='$v' order by dept3";       
      $resultc=mysql_query($queryc);  
      if (!$resultc) {
             $o->alert("69 hy_dept error!!");   
      } else {    
          $level3 = '<select name="level3no" id="level3no" onchange="xajax_showLevel4(document.getElementById(\'level3no\').value)">';  
          $level3 .= "<option value=''>請選</option>"  ;  
          while ($rowc=mysql_fetch_array($resultc)){
            $level3 .= "<option value='" . $rowc["dept3"]  . "'>" . $rowc["dept3"]."</option>";  
          }                                                                                     
          $o->assign('level3' ,"innerHTML",  $level3);                       
      }
      $o->assign('level4' ,"innerHTML", '');   
      $o->assign('level5' ,"innerHTML", ''); 
      $o->assign('level6' ,"innerHTML", '');    
      return $o;
  }
  
  Function showLevel4($v){      
      $o=new xajaxResponse();                          
      $queryc="select distinct dept4 from hy_dept where dept3='$v' order by dept4";    
      $resultc=mysql_query($queryc);  
      if (!$resultc) {
             $o->alert("69 hy_dept error!!");   
      } else {    
          $level4 = '<select name="level4no" id="level4no" onchange="xajax_showLevel5(document.getElementById(\'level4no\').value)">'; 
          $level4 .= "<option value=''>請選</option>"  ;   
          while ($rowc=mysql_fetch_array($resultc)){
            $level4 .= "<option value=" . $rowc["dept4"]  . ">" . $rowc["dept4"]."</option>";  
          }                                                                                     
          $o->assign('level4' ,"innerHTML",  $level4);                       
      }
      $o->assign('level5' ,"innerHTML", '');  
      $o->assign('level6' ,"innerHTML", '');  
      return $o;
  }
  
  Function showLevel5($v){      
      $o=new xajaxResponse(); 
      $queryc="select distinct dept5 from hy_dept where dept4='$v' order by dept5";      
      $resultc=mysql_query($queryc);  
      if (!$resultc) {
             $o->alert("69 hy_dept error!!");   
      } else {    
          $level5 = '<select name="level5no" id="level5no" onchange="xajax_showLevel6(document.getElementById(\'level5no\').value)">'; 
          $level5 .= "<option value=''>請選</option>"  ; 
          while ($rowc=mysql_fetch_array($resultc)){
            $level5 .= "<option value=" . $rowc["dept5"] . ">" . $rowc["dept5"]."</option>"; 
          }                                                                                     
          $o->assign('level5' ,"innerHTML",  $level5);                       
      }     
      $o->assign('level6' ,"innerHTML", '');                                    
      return $o;
  }      
  
  Function showLevel6($v){      
      $o=new xajaxResponse(); 
      $queryc="select distinct id, name from hy_card where dept5='$v' and quitdate='' order by id";      
      $resultc=mysql_query($queryc);  
      if (!$resultc) {
             $o->alert("69 hy_dept error!!");   
      } else {    
          $level6 = '<select name="level6no" id="level6no"'; 
          $level6 .= "<option value=''>全部</option>"  ; 
          while ($rowc=mysql_fetch_array($resultc)){
            $level6 .= "<option value=" . $rowc["id"] . ">" . $rowc["name"]."</option>"; 
          }                                                                                     
          $o->assign('level6' ,"innerHTML",  $level6);                       
      }                                                                        
      return $o;
  }   
  

                                                  
  $xajax->register(XAJAX_FUNCTION,'showLevel2');    
  $xajax->register(XAJAX_FUNCTION,'showLevel3'); 
  $xajax->register(XAJAX_FUNCTION,'showLevel4'); 
  $xajax->register(XAJAX_FUNCTION,'showLevel5');              
  $xajax->register(XAJAX_FUNCTION,'showLevel6');  
  $xajax->processRequest();

  echo '<?xml version="1.0" encoding="UTF-8"?>'; 
  $IsAjax = True; 
  
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
<p>加班費試算. </p>
<form action="<?=$PHP_SELF;?>" method="get" name="form2">
  <div align="center">
    <table width="100%"  border="0" cellpadding="4" cellspacing="1" style="border: 1px solid #cccccc">
      <tr>
        <td bgcolor="#eeeeee"><div align="left">
            月薪:
            <input name='salary' type='text' id='salary' value='<?=$salary;?>'>
            起訖日期:   
            <input name="bdate" type="text" id="bdate" size="12" maxlength="12" value=<?=$bdate;?> onfocus="new WdatePicker()"> ~~ 
            <input name="edate" type="text" id="edate" size="12" maxlength="12" value=<?=$edate;?> onfocus="new WdatePicker()"> 
            單位:   
            <span id="level1">
                <select name="level1no" id="level1no" onchange="xajax_showLevel2(document.getElementById('level1no').value)"> 
                    <option value=''>請選</option>
                    <?
                        $queryca = "select distinct dept1 from hy_dept order by dept1";
                        $resultca = mysql_query($queryca) or die ("237 hy_dept error!!");
                        if (mysql_num_rows($resultca) > 0) {
                            while ($rowca= mysql_fetch_array($resultca)) {
                              echo "<option value=" . $rowca["dept1"];
                              echo ">" . $rowca["dept1"]."</option>";
                            }
                        }
                    ?>
                </select>   
              </span>  
              <span id="level2"></span> 
              <span id="level3"></span> 
              <span id="level4"></span> 
              <span id="level5"></span>  
              <span id="level6"></span> 
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
    <th>部門</th>     
    <th>工號</th>        
    <th>姓名</th> 
    <th>日期</th>  
    <th>刷卡時間1</th>                                                                                   
    <th>刷卡時間2</th>    
    <th>刷卡時間3</th>    
    <th>刷卡時間4</th>    
    <th>加班時數</th>    
    <th>半小時不計</th> 
    <th>金額1</th> 
    <th>一小時不計</th> 
    <th>金2額</th> 
    <th>半小時計</th> 
    <th>金額3</th> 
    <th>一小時計</th> 
    <th>金額4</th> 
    <th>法定假日</th> 
    <th>金額5</th>     
    <th>合計1</th>  
    <th>合計2</th>  
    <th>合計3</th>  
    <th>合計4</th>  
  </tr>
  <?  
      $bdate=$_GET['bdate'];
      $edate=$_GET['edate'];
      if ($_GET['level6no']=='') {
        $levelfilter=" and dept5='" . $_GET['level5no'] . "' ";
      } else {
        $levelfilter=" and id='" . $_GET['level6no'] . "' ";
      } 
      
      $mm=0;
      $tmm=0; 
      
      $th=0;
      $th1=0;
      $th2=0;
      $th3=0;
      $th4=0;
      $th5=0;
      
      $tm1=0;
      $tm2=0;
      $tm3=0;
      $tm4=0;
      $tm5=0;
      
      $query = "select * from hy_card where quitdate='' and cdate2 >= '$bdate' and cdate2 <='$edate' $levelfilter  order by dept5, id, cdate2";
      $result = mysql_query($query) or die ("237 hy_card error!!" . mysql_error(). '-' . $query);  
      while ($row= mysql_fetch_array($result)) { 
       
          $mm1 = round($salary / 22 / 8 / 2 * $row['morehours1'],2);
          $mm2 = round($salary / 22 / 8  * $row['morehours2']);
          $mm3 = round($salary / 22 / 8 / 2 * $row['morehours3']);
          $mm4 = round($salary / 22 / 8  * $row['morehours4']);
          $mm5 = round($salary / 22 / 8  * $row['plushours']);
          
          $mm = $mm1+$mm2+$mm3+$mm4+$mm5;
          
          $th += $row['more1'];
          $th1 += $row['morehours1'];
          $th2 += $row['morehours2'];
          $th3 += $row['morehours3'];
          $th4 += $row['morehours4'];
          $th5 += $row['plushours'];
          $tm1 += $mm1;                                                      
          $tm2 += $mm2; 
          $tm3 += $mm3; 
          $tm4 += $mm4; 
          $tm5 += $mm5; 
                        
      
          ?>   
	        <tr bgcolor="#FFFFFF"> 
		          <td><img src="i/arrow.gif" width="16" height="16"> </td>   
			        <td><?=$row["dept5"];?></td>
              <td><?=$row["id"];?></td>   
              <td><?=$row["name"];?></td> 
              <td><?=$row["cdate2"];?></td> 
              <td><?=$row["time11"];?></td>
              <td><?=$row["time12"];?></td> 
              <td><?=$row["time21"];?></td> 
              <td><?=$row["time22"];?></td> 
              <td><?=$row["more1"];?></td> 
              <td><?=$row["morehours1"];?></td> 
              <td><?=$mm1;?></td>
              <td><?=$row["morehours2"];?></td> 
              <td><?=$mm2;?></td>
              <td><?=$row["morehours3"];?></td> 
              <td><?=$mm3;?></td>
              <td><?=$row["morehours4"];?></td>                             
              <td><?=$mm4;?></td>
              <td><?=$row["plushours"];?></td>                             
              <td><?=$mm5;?></td>  
              <td><?=$mm1+$mm5;?></td>  
              <td><?=$mm2+$mm5;?></td>  
              <td><?=$mm3+$mm5;?></td>  
              <td><?=$mm4+$mm5;?></td>  
          </tr> 
      <?  
      }   
      ?>    
  <tr>
    <td colspan='9'>&nbsp;</td>          
    <td><?=$th;?></td>       
    <td><?=$th1;?></td> 
    <td><?=$tm1;?></td> 
    <td><?=$th2;?></td> 
    <td><?=$tm2;?></td> 
    <td><?=$th3;?></td> 
    <td><?=$tm3;?></td> 
    <td><?=$th4;?></td> 
    <td><?=$tm4;?></td> 
    <td><?=$th5;?></td> 
    <td><?=$tm5;?></td>    
    <td><?=$tm1+$tm5;?></td>
    <td><?=$tm2+$tm5;?></td>
    <td><?=$tm3+$tm5;?></td>
    <td><?=$tm4+$tm5;?></td>
  </tr>
  <tr>
    <th width="16">&nbsp;</th>
    <th>部門</th>     
    <th>工號</th>        
    <th>姓名</th> 
    <th>日期</th>  
    <th>刷卡時間1</th>                                                                                   
    <th>刷卡時間2</th>    
    <th>刷卡時間3</th>    
    <th>刷卡時間4</th>    
    <th>加班時數</th>    
    <th>半小時不計</th> 
    <th>金額1</th> 
    <th>一小時不計</th> 
    <th>金額2</th> 
    <th>半小時計</th> 
    <th>金額3</th> 
    <th>一小時計</th> 
    <th>金額4</th> 
    <th>法定假日</th> 
    <th>金額5</th>    
    <th>合計1</th>  
    <th>合計2</th>  
    <th>合計3</th>  
    <th>合計4</th>   
  </tr>
  
</table>  
</form>
