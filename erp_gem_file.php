<?php
  session_start();
  $pagetitle = "系統設定 &raquo; 部門資料";
  include("_data.php");
  auth("erp_gem_file.php");
  
  if ($_POST["action"] == "update") {
       //由Oracle來查MySQL
      $sql1="select * from gem_file order by gem01";
      $erp_sql1 = oci_parse($erp_conn,$sql1 );
      oci_execute($erp_sql1);  
      while ($erp_row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) { 
          $query1 = "select gem01 from gem_file where gem01='" . $erp_row1["GEM01"] ."' limit 1";  
          $result1=mysql_query($query1, $conn) or die ('13 gem_file error!!' . mysql_error());
          if (mysql_num_rows($result1)==1){ //有找到 做更新動作
              $query2="update gem_file set        
                      gem02       = '" . $erp_row1["GEM02"]                                   . "',  
                      gem03       = '" . $erp_row1["GEM03"]                                   . "', 
                      gem04       = '" . $erp_row1["GEM04"]                                   . "', 
                      gem05       = '" . $erp_row1["GEM05"]                                   . "', 
                      gem07       = '" . $erp_row1["GEM07"]                                   . "', 
                      gemacti     = '" . $erp_row1["GEMACTI"]                                 . "', 
                      gemuser     = '" . $erp_row1["GEMUSER"]                                 . "', 
                      gemgrup     = '" . $erp_row1["GEMGRUP"]                                 . "', 
                      gemmodu     = '" . $erp_row1["GEMMODU"]                                 . "', 
                      gemdate     = '" . date('Y-m-d',strtotime($erp_row1['GEMDATE']))        . "', 
                      gem09       = '" . $erp_row1["GEM09"]                                   . "', 
                      gem10       = '" . $erp_row1["GEM10"]                                   . "', 
                      gemorig     = '" . $erp_row1["GEMORIG"]                                 . "', 
                      gemoriu     = '" . $erp_row1["GEMORIU"]                                 . "'
                      where gem01 = '" . $erp_row1["GEM01"]                                   . "' limit 1";  
          } else {
              $query2="insert into gem_file (gem01, gem02, gem03, gem04, gem05, gem07, gemacti, gemuser, gemgrup, gemmodu, gemdate, gem09, gem10, gemorig, gemoriu) values (  
                      '" . $erp_row1["GEM01"]                                 . "',   
                      '" . $erp_row1["GEM02"]                                 . "',      
                      '" . $erp_row1["GEM03"]                                 . "',     
                      '" . $erp_row1["GEM04"]                                 . "',     
                      '" . $erp_row1["GEM05"]                                 . "',     
                      '" . $erp_row1["GEM07"]                                 . "',     
                      '" . $erp_row1["GEMACTI"]                               . "',     
                      '" . $erp_row1["GEMUSER"]                               . "',     
                      '" . $erp_row1["GEMGRUP"]                               . "',     
                      '" . $erp_row1["GEMMODU"]                               . "',     
                      '" . date('Y-m-d',strtotime($erp_row1['GEMDATE']))      . "',     
                      '" . $erp_row1["GEM09"]                                 . "',     
                      '" . $erp_row1["GEM10"]                                 . "',     
                      '" . $erp_row1["GEMORIG"]                               . "',     
                      '" . $erp_row1["GEMORIU"]                               . "')"; 
          } 
          $result2=mysql_query($query2, $conn) or die ('19 gem_file update error!!'. mysql_error()); 
      }
      
      //由MySQL來查Oracle
      $query1="select gem01 from gem_file";
      $result1=mysql_query($query1, $conn);
      while ($row1=mysql_fetch_array($result1)){
          $sql1="select gem01 from gem_file where gem01='" . $row1['gem01'] ."'";
          $erp_sql1 = oci_parse($erp_conn,$sql1);
          oci_execute($erp_sql1);  
          $erp_row1=oci_fetch_array($erp_sql1, OCI_ASSOC);
          if (!$erp_row1){         //erp_row1若無資料會傳回false
              $query2="delete from gem_file where gem01='". $row1['gem01'] ."' limit 1";
              $result2=mysql_query($query2, $conn) or die ('32 gem_file delete error!!'.mysql_error());            
          }
      }       
      
      msg('資料更新完畢.');           
      forward('gem_file.php'); 
  }
  include("_header.php");
?>
<link href="css.css" rel="stylesheet" type="text/css">
<p>以下全部的部門資料!! </p>   
<form name="form1" method="post" action="<?=$PHP_SELF;?>">
    <table width="100%"  border="0" cellpadding="2" cellspacing="1" class="table4">
        <tr>
            <th width="16">&nbsp;</th>                
            <th><a href="<?=$PHP_SELF;?>?order=gem01">編號</th>   
            <th><a href="<?=$PHP_SELF;?>?order=gem02">名稱</th>          
            <th>全稱</th>
            <th>會計部門</th>  
            <th>費用類別</th>  
            <th>有效碼</th> 
            <th>所有者</th>  
            <th>所有部門</th>  
            <th>修改者</th>  
            <th>修改日</th>  
            <th>管理類別</th>  
            <th>成本中心</th>  
            <th>建立者</th>   
            <th>建立部門</th>  
        </tr>
        <?
          if(empty($_GET["order"])) $_GET["order"] = "gem01";  
          $query = "select * from gem_file order by " . $_GET["order"];
                                                                                                      
	        $result = mysql_query($query) or die ('37 gem_file error!!'. mysql_error());   
          $result=mysql_query($query);   
	        while ($row= mysql_fetch_array($result)) {
	  	        $bgcolor = "ffffff";
        ?>
	            <tr bgcolor="#<?=$bgcolor;?>">
		              <td><img src="i/arrow.gif" width="16" height="16"></td>
                  <td><?=$row["gem01"];?></td>  
		              <td><?=$row["gem02"];?></td>
		              <td><?=$row["gem03"];?></td>
		              <td><?=$row["gem05"];?></td>
                  <td><?=$row["gem07"];?></td>    
                  <td><?=$row["gemacti"];?></td>      
				          <td><?=$row["gemuser"];?></td>
                  <td><?=$row["gemgrup"];?></td> 
                  <td><?=$row["gemmodu"];?></td> 
                  <td><?=$row["gemdate"];?></td> 
                  <td><?=$row["gem09"];?></td> 
                  <td><?=$row["gem10"];?></td> 
                  <td><?=$row["gemorig"];?></td> 
                  <td><?=$row["gemoriu"];?></td> 
              </tr>
		      <?
			    }   
			    ?>
          <tr>
            <td bgcolor="#FF66FF">&nbsp;</td>
            <td><input type="hidden" name="action" value="update">
                <input type="submit" name="Submit" value="更新">          
            </td>
          </tr>
    </table>   
</form>
