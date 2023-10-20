<?
  session_start();
  $pagtitle = "IT &raquo; 把tc_lod中的金屬重量/價錢寫到tc_lof中"; 
  include("_data.php");
  include("_erp.php"); 
  
  $tc_lod005old='';
  $tc_lod001=array();
  $tc_lod002=array();
  $tc_lod003=array();
  $tc_lod004=array();
  $tc_lod005=array();
  $tc_lod006=array();
  $tc_lod007=array();
  $tc_lod008=array();
  $tc_lod009=array();
  $tc_lod010=array();
  $tc_lod011=array();
  $tc_lod012=array();  
  $tc_lod013=array();      
  
  $i=0 ;
  $ii=0;
  //取出新的料件品名
  $stclod="select tc_lod001, tc_lod002, to_char(tc_lod003,'yymmdd') tc_lod003, tc_lod004, tc_lod005, tc_lod006, tc_lod007, tc_lod008, tc_lod009, tc_lod010, tc_lod011, tc_lod012, tc_lod013
           from tc_lod_file where tc_lod003 between to_date('110301','yy/mm/dd') and to_date('110331','yy/mm/dd')  order by tc_lod004,tc_lod003,tc_lod005 "; 
  $erp_sqltclod = oci_parse($erp_conn1,$stclod);
  oci_execute($erp_sqltclod);                       
  while ($rowtclod = oci_fetch_array($erp_sqltclod, OCI_ASSOC)) {    
      $tc_lod005new=$rowtclod['TC_LOD005'];
      if ($tc_lod005new != $tc_lod005old){
          if ($tc_lod005old!='') {  //第一筆不用寫入資料 
              //此時已把同一個rX#的資料放到陣列中了             
              //先取出全部金屬的價錢
              $metal=0;
              for ($x=0; $x<$i; $x++){
                  if ((substr($tc_lod006[$x],0,1)=='M') or ($tc_lod006[$x]=='O002')) {
                      $metal+=$tc_lod010[$x];
                  }
              }
              //將金屬價格加到第一個產品身上
              for ($x=0; $x<$i; $x++){
                  if ((substr($tc_lod006[$x],0,1)!='M') and ($tc_lod006[$x]!='O002')) {   
                      $ii++;
                      if ($metal==0) {
                          $tc_loe001=$tc_lod001[$x];
                          $tc_loe002=$tc_lod002[$x];     
                          $tc_loe003=$tc_lod003[$x];     
                          $tc_loe004=$tc_lod004[$x];     
                          $tc_loe005=str_replace("'","",$tc_lod005[$x]);     
                          $tc_loe006=$tc_lod006[$x];     
                          $tc_loe007=$tc_lod007[$x];     
                          $tc_loe008=$tc_lod008[$x];     
                          $tc_loe009=$tc_lod009[$x];     
                          $tc_loe010=$tc_lod010[$x];     
                          $tc_loe011=$tc_lod011[$x];     
                          $tc_loe012=$tc_lod012[$x];     
                          $tc_loe013=$tc_lod013[$x];     
                          $tc_loe014=0;  
                      } else {
                          $tc_loe001=$tc_lod001[$x];
                          $tc_loe002=$tc_lod002[$x];     
                          $tc_loe003=$tc_lod003[$x];     
                          $tc_loe004=$tc_lod004[$x];     
                          $tc_loe005=str_replace("'","",$tc_lod005[$x]);     
                          $tc_loe006=$tc_lod006[$x];     
                          $tc_loe007=$tc_lod007[$x];     
                          $tc_loe008=$tc_lod008[$x];  
                          $tc_loe010=$tc_lod010[$x]+$metal;    
                          $tc_loe009=$tc_loe010/$tc_loe008;  
                          $tc_loe011=$tc_lod011[$x];     
                          $tc_loe012=$tc_lod012[$x];     
                          $tc_loe013=$tc_lod013[$x];     
                          $tc_loe014=$metal;   
                          $metal=0;                                  
                      } 
                      $stcloe="insert into tc_lof_file values ('$tc_loe001',$ii, to_date('$tc_loe003','yy/mm/dd'),'$tc_loe004','$tc_loe005','$tc_loe006','$tc_loe007',$tc_loe008,$tc_loe009,$tc_loe010,'$tc_loe011','$tc_loe012','$tc_loe013','','','','','','','','','',$tc_loe014  ) "; 
                      $erp_sqltcloe = oci_parse($erp_conn1,$stcloe );
                      oci_execute($erp_sqltcloe); 
                  }  
              } 
              //儲存完後 把陣列清空
              $tc_lod001=array();
              $tc_lod002=array();
              $tc_lod003=array();
              $tc_lod004=array();
              $tc_lod005=array();
              $tc_lod006=array();
              $tc_lod007=array();
              $tc_lod008=array();
              $tc_lod009=array();
              $tc_lod010=array();
              $tc_lod011=array();
              $tc_lod012=array();  
              $tc_lod013=array();   
              
              $i=0;    
          }             
      } 
      $tc_lod005old=$rowtclod['TC_LOD005'];
      $tc_lod001[$i]=$rowtclod['TC_LOD001'] ;
      $tc_lod002[$i]=$rowtclod['TC_LOD002'] ;
      $tc_lod003[$i]=$rowtclod['TC_LOD003'] ;
      $tc_lod004[$i]=$rowtclod['TC_LOD004'] ;
      $tc_lod005[$i]=$rowtclod['TC_LOD005'] ;
      $tc_lod006[$i]=$rowtclod['TC_LOD006'] ;
      $tc_lod007[$i]=$rowtclod['TC_LOD007'] ;
      $tc_lod008[$i]=floatval($rowtclod['TC_LOD008']) ;
      $tc_lod009[$i]=floatval($rowtclod['TC_LOD009']) ;
      $tc_lod010[$i]=floatval($rowtclod['TC_LOD010']) ;
      $tc_lod011[$i]=$rowtclod['TC_LOD011'] ;
      $tc_lod012[$i]=$rowtclod['TC_LOD012'] ;
      $tc_lod013[$i]=$rowtclod['TC_LOD013'] ;  
      $i++;                                    
               
  }
  
  msg('Done!!');
?>    
