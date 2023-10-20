<?php
  session_start();
  $pagetitle = "業務部 &raquo; 出貨統計V7";
  include("_data.php");
  include("_erp.php");
  date_default_timezone_set('Asia/Taipei');      
  
  $bdate = date('Y-m-d');                         

  
  //產生email內容的html檔      
  $emailbody=file_get_contents('header_out.html');
  $a811total=0;
  $a821total=0;
  $a831total=0;
  $a8total=0;


  $query="select occ02, a811, a821, a831, a811+a821+a831 a8 from erp_dailycount where qtype='U' and iotype='OU' and tdate='$bdate' order by occ02 ";
  $result = mysql_query($query) or die ('393 erp_dailycount read error!!' . mysql_error()); 

  while ($row= mysql_fetch_array($result)) {  
       $emailbody .= '<tr>';
       $emailbody .= '<td class="col0">' . $row['occ02'] . '</td>';
       if ($row['a811']>0 ) {
          $emailbody .= '<td class="col1">' . $row['a811'] . '</td>';
       } else {
          $emailbody .= '<td class="col1">&nbsp;</td>';
       }  
       if ($row['a821']>0 ) {
          $emailbody .= '<td class="col2">' . $row['a821'] . '</td>';
       } else {
          $emailbody .= '<td class="col2">&nbsp;</td>';
       }  
       if ($row['a831']>0 ) {
          $emailbody .= '<td class="col3">' . $row['a831'] . '</td>';
       } else {
          $emailbody .= '<td class="col3">&nbsp;</td>';
       }  
       $emailbody .= '<td class="col4">' . $row['a8'] . '</td>';      
       $emailbody .= '</tr>'; 
       
       $a811total += $row['a811'];
       $a821total += $row['a821'];
       $a831total += $row['a831'];
       $a8total += $row['a8'];             
      
  }

  $emailbody .= '<tr>';              
  $emailbody .= '<td class="col0">合計</td>';    
  $emailbody .= '<td class="col1">' . $a811total . '</td>';
  $emailbody .= '<td class="col2">' . $a821total . '</td>';
  $emailbody .= '<td class="col3">' . $a831total . '</td>';
  $emailbody .= '<td class="col4">' . $a8total . '</td>';
  $emailbody .= '</tr>';       
  $emailbody .= file_get_contents('footer_out.html'); 
  
  $filename='email/' . $bdate . '_VD_DailyCasesDelivered.xls';
    
  require 'PHPMailer/PHPMailerAutoload.php';

  $mail = new PHPMailer();   
  $mail->isSMTP();
  //Enable SMTP debugging
  // 0 = off (for production use)
  // 1 = client messages
  // 2 = client and server messages
  $mail->SMTPDebug = 0;
  //Ask for HTML-friendly debug output
  $mail->Debugoutput = 'html';
  //Set the hostname of the mail server   

  $_SESSION['emailserver']='smtp.vedendentalgroup.com';
  $_SESSION['emailport'] = 25;    
  $_SESSION['emailauth'] = true;
  $_SESSION['emailusername'] ='frank@vedendentalgroup.com';
  $_SESSION['emailpassword'] ='Aj8900!@#'; 

  $mail->CharSet = 'utf-8';     
  $mail->Host = $_SESSION['emailserver'];                  
  $mail->Port = $_SESSION['emailport'];  
  $mail->SMTPAuth = $_SESSION['emailauth'];  
  $mail->Username = $_SESSION['emailusername'];  
  $mail->Password = $_SESSION['emailpassword'];
  //Set who the message is to be sent from
  $mail->setFrom('frank@vedendentalgroup.com', 'Frank Yu'); 
  $mail->addReplyTo('frank@vedendentalgroup.com.com', 'Frank Yu');  
  //$mail->addAddress('it@vedendentalgroup.com', 'Veden');     
  $mail->addAddress('casesreceived@vedendentalgroup.com', 'Veden'); 
  $mail->addAddress('casesreceived@jordanlabo.com', 'Jordan'); 
  $mail->addAddress('wenhsien.cheng@gmail.com', 'Mr. BIG'); 
  $mail->Subject = 'Veden '. $bdate .' 出貨統計 -- 依客戶別';
  //$mail->Body    = $_SESSION['emailbody'];

  $mail->addAttachment($filename, $filename);  

  //file_put_contents('1.html',$emailbody);
  //$mail->msgHTML(file_get_contents('1.html'), dirname(__FILE__)); 

  $mail->msgHTML($emailbody);
  //$mail->AltBody = 'This is a plain-text message body';   
  //$mail->addAttachment('images/phpmailer_mini.gif');
  $mail->addEmbeddedImage('pies/' . $bdate . '_out.jpg','pie', $bdate.'_out.jpg');
  $mail->send();
?>  
        
                            