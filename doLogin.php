<? 
      session_start();
      include_once("_data.php"); 

      if (($_POST['account']) && ($_POST['pass'])) { 
	      $time=time(); 
          $account = safetext($_POST["account"]);
          $pass = safetext($_POST["pass"]);
          $host = safetext($_SERVER["HTTP_HOST"]);
          $query = "select guid, account, cname, password, maker, lastlogin from users where account = '" . $account . "' AND password = '" . md5($pass) . "' limit 1";
          $result = mysql_query($query) or die ( '11 Users error!!');
          if (mysql_num_rows($result) == 0) {
              // 20230509RAYMOND_temp
              // $querya = "insert into accesslogs values ('" . uuid(). "','" . $user ."','". $pass  ."','index.php','N','".$_SERVER['REMOTE_ADDR']."' , '" . safetext($_SERVER['REQUEST_URI']) . "' , '', '" . safetext($_SERVER['HTTP_REFERER']) . "' , NULL)";             
              // $resulta = mysql_query($querya) or die ( '14 AccessLog error!!');
              Header("Location: index.php");
          } else {
              // 20230509RAYMOND_temp
              // $querya = "insert into accesslogs values ('" . uuid(). "','" . $user ."','". $pass  ."','index.php','Y','".$_SERVER['REMOTE_ADDR']."' , '" . safetext($_SERVER['REQUEST_URI']) . "' , '', '" . safetext($_SERVER['HTTP_REFERER']) . "' , NULL)";        
              // $resulta = mysql_query($querya) or die ( '20 AccessLogs error!!');
              $row = mysql_fetch_array($result);
              //if ($_POST["cookie"] == "1") {
              //    setcookie ("account", $row['account'], $time+999999); 
              //    setcookie ("password", $row['password'], $time+999999);               
              //} else {
              //    setcookie ("account", $row['account'], $time-100); 
              //    setcookie ("password", $row['password'], $time-100);     
              //} 
              $_SESSION['account']      = $row['account'];
              $_SESSION['name']         = $row["cname"];   
              $_SESSION['userguid']     = $row["guid"];  
              $_SESSION['maker']        = $row["maker"];      
              $_SESSION['lastlogin']    = $row["lastlogin"];
              $query = "update users set lastlogin = CURRENT_TIMESTAMP where guid = '" . safetext($row["guid"]) . "' limit 1";
              $result = mysql_query($query) or die ('32 Users updated error!! ');    
              Header("Location: main.php");   
          }    
      } else {
        Header("Location: index.php");
      }
?>