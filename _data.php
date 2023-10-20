<? 
    $db_host = "localhost:3302";
    $db_user = "root";
    $db_pass = "veden*()";
    $db_name = "erpdb";   
    // make the connection
    $conn = mysql_connect($db_host,$db_user,$db_pass) or die('Oracle 11G server connect error!!');
    mysql_query("SET NAMES 'utf8'");
    mysql_select_db($db_name,$conn) or die('Database connect error!!'); 
    $_SESSION["systemname"] = "Web ERP管理系統";
    
          //為方便移值到其他公司 在此定義各公司使用的資料庫名字
    // $outserver: 境外  $inserver:境內
    $_SESSION['outserver']='vd210';
    $_SESSION['inserver']='vd110';
    
    $erp_db_host = "topprod";
    $erp_db_user = "vd110";
    $erp_db_pass = "vd110";  
    $erp_conn = oci_connect($erp_db_user, $erp_db_pass, $erp_db_host,'AL32UTF8');
    
    $erp_db_host1 = "topprod";
    $erp_db_user1 = "vd110";
    $erp_db_pass1 = "vd110";  
    $erp_conn1 = oci_connect($erp_db_user1, $erp_db_pass1, $erp_db_host1,'AL32UTF8'); 
     
    $erp_db_host2 = "topprod";
    $erp_db_user2 = "vd210";
    $erp_db_pass2 = "vd210";  
    $erp_conn2 = oci_connect($erp_db_user2, $erp_db_pass2, $erp_db_host2,'AL32UTF8'); 

    $_SESSION['emailserver']='smtp.exmail.qq.com';
    $_SESSION['emailport'] = 25;    
    $_SESSION['emailauth'] = true;
    $_SESSION['emailusername'] ='report@veden.dental';
    $_SESSION['emailpassword'] ='Veden@123';
    
function safetext($data) {
    $data = mysql_real_escape_string($data);
    $data = htmlspecialchars($data);
    return $data;
}

function msg($msg) {
        echo "<script language=javascript>
        <!--
        alert('" . $msg . "');
        //-->
        </script>";
}

function forward($forward) {
        echo "<script language=javascript>
        <!--
        location.href = '" . $forward . "';
        //-->
        </script>";
}    

function forward_new($forward) {
        echo "<script language=javascript>
        <!--
        location.href = '" . $forward . "' target=_blank;
        //-->
        </script>";
}  

function auth_old($filename) {
   //改寫要檢查在aps, aprights 裡面有無設定資料
   $query = "select r.apguid rapguid, r.userguid ruserguid, a.guid aguid, a.filename afilename from aprights r, aps a " .
            " where r.apguid = a.guid and a.filename='" .$filename . "' and r.userguid='" . $_SESSION["userguid"] . "' limit 1";
   $result = mysql_query($query) or die ('39 APRights error.');
   if (mysql_num_rows($result) == 0) {
       // 20230509RAYMOND_temp
     //$querya = "insert into accesslogs (guid, account, password, ap, granted, ip ) values " .
//               "('" . uuid(). "','" . safetext($_SESSION['account']) ."','','" . safetext($filename) . "','N','".$_SERVER['REMOTE_ADDR']."')";
//     $resulta = mysql_query($querya) or die ( '41 AccessLog error!!');
     msg("權限不足.");
     forward("doLogout.php");
   }else{
       // 20230509RAYMOND_temp
//     $querya = "insert into accesslogs (guid, account, password, ap, granted, ip ) values " .   
//               "('" . uuid(). "','" . safetext($_SESSION['accpimt']) ."','','" . safetext($filename) . "','Y','".$_SERVER['REMOTE_ADDR']."')";   
//     $resulta = mysql_query($querya) or die ( '46 AccessLog error!!');
   } 
}

function auth($filename) {
   //改寫要檢查在aps, aprights 裡面有無設定資料
   
   if (is_null($_SESSION['userguid']) || $_SESSION['userguid']==''|| is_null($_SESSION['account'] || $_SESSION['account']=='') ){
       // 20230509RAYMOND_temp
//       $querya = "insert into accesslogs values ('" . uuid(). "','" . safetext($_SESSION['account']) ."','','" . safetext($filename) . "','Y','".$_SERVER['REMOTE_ADDR']."' , '" . substr(safetext($_SERVER['REQUEST_URI']),0,255) ."' , '" . $_SESSION["userguid"] . "' , '" . substr(safetext($_SERVER['HTTP_REFERER']),0,255) . "' , NULL)"; 
       die();
   } else {
       if ($_SESSION['account']=='7a5916fcf43955da8badda67b6ca3187'){           
       } else {
           $query = "select r.apguid rapguid, r.userguid ruserguid, a.guid aguid, a.filename afilename from aprights r, aps a " .
                    " where r.apguid = a.guid and a.filename='" .$filename . "' and r.userguid='" . $_SESSION["userguid"] . "' limit 1";
           $result = mysql_query($query) or die ('APRighs connect error.');
           if (mysql_num_rows($result) == 0) {
               // 20230509RAYMOND_temp
//             $querya = "insert into accesslogs values ('" . uuid(). "','" . safetext($_SESSION['account']) ."','','" . safetext($filename) . "','N','".$_SERVER['REMOTE_ADDR']."' , '" . substr(safetext($_SERVER['REQUEST_URI']),0,255) ."' , '" . $_SESSION["userguid"] . "' , '" . substr(safetext($_SERVER['HTTP_REFERER']),0,255) . "' , NULL)";    
//             $resulta = mysql_query($querya) or die ( '41 Access error!!'.mysql_error());
             msg("Access denied.");
             forward("doLogout.php");
           }else{
               // 20230509RAYMOND_temp
//             $querya = "insert into accesslogs values ('" . uuid(). "','" . safetext($_SESSION['account']) ."','','" . safetext($filename) . "','Y','".$_SERVER['REMOTE_ADDR']."' , '" . substr(safetext($_SERVER['REQUEST_URI']),0,255) ."' , '" . $_SESSION["userguid"] . "' , '" . substr(safetext($_SERVER['HTTP_REFERER']),0,255) . "' , NULL)";    
//             $resulta = mysql_query($querya) or die ( '46 Access error!!'.mysql_error());
           } 
       }
   }
}


function getClientDetails () {
    $query = "select * from clients where guid = '" . $_SESSION["cguid"] . "' limit 1";
    $result = mysql_query($query) or die ('Client not found.');
    if (mysql_num_rows($result) > 0) {
        $clientrow = mysql_fetch_array($result);
    }else{
        die('Client not found.');
    }
    return $clientrow;
}

function generateOrderNo($cguid) {
      $query = "select id from clients where guid='" . $cguid . "' limit 1";
      $result=mysql_query($query) or die ('87 Clients error!!');
      if (mysql_num_rows($result) > 0) {  
          $row=mysql_fetch_row($result);
          $clientno=$row[0];

          $query = "select no from ordersno where clientguid='" . $cguid . "' limit 1 ";
          $result = mysql_query($query);
          if (mysql_num_rows($result) > 0) {
                $row = mysql_fetch_row($result); 
                $rowno= $row[0];
                if ( substr($rowno,0,4) == strval(date('Y'))) {
                    $rowno=strval(intval($rowno)+1);
                } else {
                    $rowno=strval(date('Y')) . "000001"; 
                }                    
                $query="update ordersno set no = '". $rowno . "',
                                             timestamp = NULL 
                                             where clientguid ='" . $cguid . "'";
                $result= mysql_query($query) or die("107 OrdersNo error!!"); 
          } else {
            $rowno=strval(date('Y')) . "000001";
            $query="insert into ordersno value ('". $cguid . "','" . $rowno . "',NULL)";
            $result= mysql_query($query) or die("111 OrdersNo error!!");     
          }
          $returnvalue = $clientno . $rowno; 
          $query = "insert into ordersnohistories value ('" . uuid() . "','" . $cguid . "','". $returnvalue . 
                   "','" . safetext($_SESSION["uname"]) . "','" . $_SERVER['REMOTE_ADDR']."' ,NULL)"; 
          $result = mysql_query($query) or die ('116 OrdersNoHistories error!!');  

          return ($returnvalue); 
          
      } else {
        die ('114 Client not found!!');
      }
}

function generateClientInvoiceNo($cguid) {
      $query = "select id from clients where guid='" . $cguid . "' limit 1";
      $result=mysql_query($query) or die ('87 Clients error!!');
      if (mysql_num_rows($result) > 0) {  
          $row=mysql_fetch_row($result);
          $clientno=$row[0];

          $query = "select no from clientsinvoiceno where clientguid='" . $cguid . "' limit 1 ";
          $result = mysql_query($query);
          if (mysql_num_rows($result) > 0) {
                $row = mysql_fetch_row($result); 
                $rowno= $row[0];
                if ( substr($rowno,0,4) == strval(date('Y'))) {
                    $rowno=strval(intval($rowno)+1);
                } else {
                    $rowno=strval(date('Y')) . "000001"; 
                }                    
                $query="update clientsinvoiceno set no = '". $rowno . "',
                                             timestamp = NULL 
                                             where clientguid ='" . $cguid . "'";
                $result= mysql_query($query) or die("InvoiceNo error!!"); 
          } else {
            $rowno=strval(date('Y')) . "000001";
            $query="insert into clientsinvoiceno value ('". $cguid . "','" . $rowno . "',NULL)";
            $result= mysql_query($query) or die("OrderNo error!!");     
          }
          $returnvalue = $clientno . $rowno; 
          $query = "insert into clientsinvoicenohistories value ('" . uuid() . "','" . $cguid . "','". $returnvalue . 
                   "','" . safetext($_SESSION["uname"]) . "','" . $_SERVER['REMOTE_ADDR']."' ,NULL)"; 
          $result = mysql_query($query) or die ('InvoiceNoHistories error!!');  

          return ($returnvalue); 
          
      } else {
        die ('114 Client not found!!');
      }
}

function showday($datum) {
    echo date('l', mktime(0, 0, 0, substr($datum, 5, 2), substr($datum, 8, 2), substr($datum, 0, 4)));
} 

function currency($data) {
    return "&euro; " . number_format($data, 2, ".", " ");
}

function uuid($serverID=1)
{
    $t=explode(" ",microtime());
    return sprintf( '%04x-%08s-%08s-%04s-%04x%04x',
        $serverID,
        clientIPToHex(),
        substr("00000000".dechex($t[1]),-8),   // get 8HEX of unixtime
        substr("0000".dechex(round($t[0]*65536)),-4), // get 4HEX of microtime
        mt_rand(0,0xffff), mt_rand(0,0xffff));
}

function uuidDecode($uuid) {
    $rez=Array();
    $u=explode("-",$uuid);
    if(is_array($u)&&count($u)==5) {
        $rez=Array(
            'serverID'=>$u[0],
            'ip'=>clientIPFromHex($u[1]),
            'unixtime'=>hexdec($u[2]),
            'micro'=>(hexdec($u[3])/65536)
        );
    }
    return $rez;
}

function clientIPToHex($ip="") {
    $hex="";
    if($ip=="") $ip=getEnv("REMOTE_ADDR");
    $part=explode('.', $ip);
    for ($i=0; $i<=count($part)-1; $i++) {
        $hex.=substr("0".dechex($part[$i]),-2);
    }
    return $hex;
}

function clientIPFromHex($hex) {
    $ip="";
    if(strlen($hex)==8) {
        $ip.=hexdec(substr($hex,0,2)).".";
        $ip.=hexdec(substr($hex,2,2)).".";
        $ip.=hexdec(substr($hex,4,2)).".";
        $ip.=hexdec(substr($hex,6,2));
    }
    return $ip;
}

class System
{
    function currentTimeMillis()
    {
        list($usec, $sec) = explode(" ", microtime());
        return $sec.substr($usec, 2, 3);
    }
}
 
class NetAddress
{
    var $Name = 'localhost';
    var $IP = '127.0.0.1';

    function getLocalHost()
    {
        $address = new NetAddress();
        $address->Name = $_ENV["COMPUTERNAME"];
        $address->IP = $_SERVER["SERVER_ADDR"];
        return $address;
    }

    function toString()
    {
        return strtolower($this->Name.'/'.$this->IP);
    }
}

class Random
{
    function nextLong()
    {
        $tmp = rand(0, 1) ? '-' : '';
        return $tmp.rand(1000, 9999).rand(1000, 9999).rand(1000, 9999).rand(100,
            999).rand(100, 999);
    }
}

class Guid
{
    var $valueBeforeMD5;
    var $valueAfterMD5;

    function Guid()
    {
        $this->getGuid();
    }

    function getGuid()
    {
        $address = NetAddress::getLocalHost();
        $this->valueBeforeMD5 = $address->toString().':'.System
            ::currentTimeMillis().':'.Random::nextLong();
        $this->valueAfterMD5 = md5($this->valueBeforeMD5);
    }

    function newGuid()
    {
        $Guid = new Guid();
        return $Guid;
    }

    function toString()
    {
        $raw = strtoupper($this->valueAfterMD5);
        return substr($raw, 0, 8).'-'.substr($raw, 8, 4).'-'.substr($raw, 12, 4).'
                      -'.substr($raw, 16, 4).'-'.substr($raw, 20);
    }
}

function utf8_substr($StrInput,$strStart,$strLen)
{
    //對字串做URL Eecode
    $StrInput = mb_substr($StrInput,$strStart,mb_strlen($StrInput));
    $iString = urlencode($StrInput);
    $lstrResult="";
    $istrLen = 0;
    $k = 0;
    do{
        $lstrChar = substr($iString, $k, 1);
        if($lstrChar == "%"){
            $ThisChr = hexdec(substr($iString, $k+1, 2));
            if($ThisChr >= 128){
                if($istrLen+3 < $strLen){
                    $lstrResult .= urldecode(substr($iString, $k, 9));
                    $k = $k + 9;
                    $istrLen+=3;
                }else{
                    $k = $k + 9;
                    $istrLen+=3;
                }
            }else{
                $lstrResult .= urldecode(substr($iString, $k, 3));
                $k = $k + 3;
                $istrLen+=2;
            }
        }else{
            $lstrResult .= urldecode(substr($iString, $k, 1));
            $k = $k + 1;
            $istrLen++;
        }
        
    // strlen仍要改寫
    }while ($k < strlen($iString) && $istrLen < $strLen); 
    return $lstrResult;
} 

function no2column($inputno){
  $allstring='-ABCDEFGHIJKLMNOPQRSTUVWXYZ-';
  $devide = intval($inputno/26);
  $rest= $inputno % 26;
  if ($devide > 0) { 
    $column = substr($allstring,$devide,1);
  } else {
    $column ='';
  }
  $column .= substr($allstring,$rest,1);
  return $column;
}

// modify by mao 2013/07/24
$mailer_from = "m@vedendentalgroup.com";
$mailer_replyto = "m@vedendentalgroup.com";
$mailer_send = "m@vedendentalgroup.com";

/*
$server_name = "www.vedendentalgroup.com";

$mail_host = 'smtp.vedendentalgroup.com';
$mail_port = '587';
$mail_username = 'scan@vedendentalgroup.com';
$mail_password = '1234@com"';
$mail_from = 'scan@vedendentalgroup.com';
$mail_reply = 'm@vedendentalgroup.com';
$mail_to = 'm@vedendentalgroup.com';
*/


// @ini_set('SMTP','smtp.vedendentalgroup.com');
// @ini_set('smtp_port','25');

// @ini_set('display_errors', 'on');
// error_reporting(E_ALL ^ E_NOTICE);
?>
