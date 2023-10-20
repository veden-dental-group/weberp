<?php
  session_start();
  $pagetitle = "報工查詢 &raquo; ";
  include("_data.php");
  auth("maker_performance.php");
  date_default_timezone_set('Asia/Taipei');   
  if ($_GET["action"] == "del") {
      //delete the user's app. rights
      $query = "delete from aprights where userguid='" . safetext($_GET["guid"]) . "'";
      $result=mysql_query($query) or die ("Aprights deleted error!!");
      //delet the user' data
      $query = "delete from users where guid = '" . safetext($_GET["guid"]) . "' limit 1";
      $result = mysql_query($query) or die ('Users deleted error!!');        
      msg('帳號已刪除.');
      forward('users.php'); 
  }
  include("_header.php");
  
  
  $maker=$_GET["maker"];    
  $s1= "select gem02 from gem_file where gem01='$maker'";
  $erp_sql1 = oci_parse($erp_conn,$s1 );
  oci_execute($erp_sql1);  
  $row1 = oci_fetch_array($erp_sql1, OCI_ASSOC);
  $makername=$row1['GEM02']; 
  //$maker='6A0000';
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p> <b><?=$makername;?></b> </p>   
<table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
    <tr>
        <th width="16">&nbsp;</th> 
        <th>製處</th>    
        <th>工序</th>          
        <th>在製顆數</th>
        <th>進站4小时</th>  
        <th>進站8小時</th>   
        <th>進站8小時以上</th>  
        <th>上站出本站未進</th>    
    </tr>                   
    <?
       // select (sysdate-to_date(to_char(tc_srg007,'yyyy-mm-dd') || tc_srg008 ,'yyyy-mm-dd hh24:mi:ss'))*24*60 from tc_srg_file where tc_srg007 is not null
                        
      //找出本製處的所有技工卡號
      //找出本技工刷進且有在製量的顆數 
      $s1="(select ecb06, ecb17, tc_srg005, 0 ins4, 0 ins8, 0 ing8, 0 outs1, 0 outs2, 0 outg2 from 
            ( select tc_srg003, tc_srg004,tc_srg005 from tc_srg_file where tc_srg005 is not null 
              and tc_srg009 in ( select gen01 from gen_file where gen03='$maker')) a, ecb_file " .
          "where ecb01=tc_srg003 and ecb03=tc_srg004) " ;
      //有刷進沒有刷出 且時間 < 4 小時
      $s2="(select ecb06, ecb17, 0 tc_srg005, tc_srg005 ins4, 0 ins8, 0 ing8, 0 outs1, 0 outs2, 0 outg2 from 
            ( select tc_srg003, tc_srg004,tc_srg005 from tc_srg_file where tc_srg005 is not null 
              and tc_srg007 is not null and tc_srg010 is null 
              and ((sysdate-to_date(to_char(tc_srg007,'yyyy-mm-dd') || tc_srg008 ,'yyyy-mm-dd hh24:mi:ss'))*24*60) < 240
              and tc_srg009 in ( select gen01 from gen_file where gen03='$maker')) a, ecb_file " .
          "where ecb01=tc_srg003 and ecb03=tc_srg004) " ;
       //有刷進沒有刷出 且時間 >= 4小時, < 8 小時
      $s3="(select ecb06, ecb17, 0 tc_srg005, 0 ins4, 0 ins8, tc_srg005 ing8, 0 outs1, 0 outs2, 0 outg2 from 
            ( select tc_srg003, tc_srg004,tc_srg005 from tc_srg_file where tc_srg005 is not null 
              and tc_srg007 is not null and tc_srg010 is null 
              and ((sysdate-to_date(to_char(tc_srg007,'yyyy-mm-dd') || tc_srg008 ,'yyyy-mm-dd hh24:mi:ss'))*24*60) >= 240
              and ((sysdate-to_date(to_char(tc_srg007,'yyyy-mm-dd') || tc_srg008 ,'yyyy-mm-dd hh24:mi:ss'))*24*60) < 480     
              and tc_srg009 in ( select gen01 from gen_file where gen03='$maker')) a, ecb_file " .
          "where ecb01=tc_srg003 and ecb03=tc_srg004) " ;   
      //有刷進沒有刷出 且時間 >= 4 小時
      $s4="(select ecb06, ecb17, 0 tc_srg005, 0 ins4, 0 ins8, tc_srg005 ing8, 0 outs1, 0 outs2, 0 outg2 from 
            ( select tc_srg003, tc_srg004,tc_srg005 from tc_srg_file where tc_srg005 is not null 
              and tc_srg007 is not null and tc_srg010 is null 
              and ((sysdate-to_date(to_char(tc_srg007,'yyyy-mm-dd') || tc_srg008 ,'yyyy-mm-dd hh24:mi:ss'))*24*60) >= 480
              and tc_srg009 in ( select gen01 from gen_file where gen03='$maker')) a, ecb_file " .
          "where ecb01=tc_srg003 and ecb03=tc_srg004) " ;           
          
      //有在製量 但無刷進記錄 表示 上一站刷出 本站卻未刷入
      $s5="(select ecb06, ecb17, 0 tc_srg005, 0 ins4, 0 ins8, 0 ing8, tc_srg005 outs1, 0 outs2, 0 outg2 from 
            ( select tc_srg003, tc_srg004,tc_srg005 from tc_srg_file where tc_srg005 is not null 
              and tc_srg007 is null 
              and tc_srg009 in ( select gen01 from gen_file where gen03='$maker')) a, ecb_file " .
          "where ecb01=tc_srg003 and ecb03=tc_srg004) " ;
              
      $sall="select ecb06, ecb17, sum(tc_srg005) tc_srg005, sum(ins4) ins4, sum(ins8) ins8, sum(ing8) ing8, sum(outs1) outs1, sum(outs2) outs2, sum(outg2) outg2 from " .
            "( " . $s1 . " union all " . $s2 . " union all " . $s3 .  " union all " . $s5 . " ) aa group by (ecb06,ecb17) order by (ecb06||ecb17)" ; 

      $erp_sql1 = oci_parse($erp_conn,$sall );
      oci_execute($erp_sql1);
      $total=0;
      while ($row = oci_fetch_array($erp_sql1, OCI_ASSOC)) {  
	  	    $bgkleur = "ffffff";
          $total+=$row["TC_SRG005"];
    ?>
	        <tr bgcolor="#<?=$bgkleur;?>">
		          <td><img src="i/arrow.gif" width="16" height="16"></td>
              <td><?=$makername;?></td>
              <td><?=$row["ECB06"];?>--<?=$row["ECB17"];?></td> 
		          <td><?=$row["TC_SRG005"];?></td>
		          <td><?=$row["INS4"];?></td>
              <td><?=$row["INS8"];?></td>      
				      <td><?=$row["ING8"];?></td> 
              <td><?=$row["OUTS1"];?></td>
          </tr>
		  <?
			}
			?>
      <tr>
            <td><img src="i/arrow.gif" width="16" height="16"></td>
            <td>合計</td>
            <td>&nbsp;</td> 
            <td><?=$total;?></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>   
            <td>&nbsp;</td>   
            <td>&nbsp;</td>          
      </tr>
      
</table>   
