<?php
include("_data.php");
include("_erp.php");
      $edate ='2017-01-31';
      $query2= "delete from stockin where countdate='$edate'";
      $result2= mysql_query($query2) or die ('1154 Stockin Deleted error. ' .mysql_error());
      //入庫 1
      $s1="select to_char(rvu03,'yyyy-mm-dd') rvu03, rvv31, ima02, rvvud02, sum(rvvud07) rvvud07 from rvu_file, rvv_file, ima_file " .
          "where rvuconf='Y' and rvu00='1' and rvu03<=to_date('$edate','yy/mm/dd') and rvu01=rvv01 and rvv31=ima01 and substr(ima06,1,1)!='9' group by rvu03,rvv31,ima02,rvvud02 ";
      $erp_sql1 = oci_parse($erp_conn1,$s1 );
      oci_execute($erp_sql1);
      while ($row1 = oci_fetch_array($erp_sql1, OCI_ASSOC)) {
          $rvu03=$row1['RVU03'];   //日期
          $rvv31=$row1['RVV31'];   //料號
          $ima02=$row1['IMA02'];   //品名
          $rvvud02=$row1['RVVUD02'];   //入庫單位
          $rvvud07=is_null($row1['RVVUD07'])?0:$row1['RVVUD07'];   //入庫量
          //check if this is a new date
          $qq="select indate from stockin where countdate='$edate' and code='$rvv31' ";
          $rq=mysql_query($qq);
          $row= mysql_fetch_array($rq);
          $indate = $row['indate'];
          if (is_null($indate)){
            $query1= "insert into stockin ( countdate, code, name, unit, indate, qty, source ) values (
                     '" . $edate      . "',
                     '" . $rvv31      . "',
                     '" . $ima02      . "',
                     '" . $rvvud02    . "',
                     '" . $rvu03      . "',
                     '" . $rvvud07    . "','1')";
            $result1= mysql_query($query1) or die ('1163 Stockin Added error. ' .mysql_error());
          } else {
            if ($rvu03>$indate){
              $query1= "update stockin set indate='$rvu03' where countdate='$edate' and code='$rvv31' ";
              $result1= mysql_query($query1) or die ('1163 Stockin Added error. ' .mysql_error());
            }
               
          }
      }
