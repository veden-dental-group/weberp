<?php
// add by mao 2013/07/24

// configuration
$server_name = "www.vedendentalgroup.com";

$mail_host = 'smtp.vedendentalgroup.com';
$mail_port = '587';
$mail_username = 'report@vedendentalgroup.com';
$mail_password = '1234@com"';
$mail_from = 'report@vedendentalgroup.com';
/*
$mail_reply = 'm@vedendentalgroup.com';
$mail_to = 'm@vedendentalgroup.com';
*/

// preparation
require_once './Classes/PHPMailer/class.phpmailer.php';

// $_SERVER['SERVER_NAME'] = 'www.vedendentalgroup.com';
$_SERVER['SERVER_NAME'] = $server_name;
// $regards = "Veden Dental Group Inc.";

$mail = new PHPMailer();
$mail->IsSMTP();
$mail->SMTPDebug  = 0;

//Set the hostname of the mail server
$mail->Host       = 'smtp.vedendentalgroup.com';
//Set the SMTP port number - 587 for authenticated TLS, a.k.a. RFC4409 SMTP submission
$mail->Port       = 587;
//Set the encryption system to use - ssl (deprecated) or tls
$mail->SMTPSecure = 'tls';

//Whether to use SMTP authentication
$mail->SMTPAuth   = true;
//Username to use for SMTP authentication - use full email address for gmail
$mail->Username   = 'report@vedendentalgroup.com';
//Password to use for SMTP authentication
$mail->Password   = '1234@com';

// mail specific data
$mail->SetFrom('report@vedendentalgroup.com', 'Veden Dental Group');
/*
$mail->AddReplyTo('m@vedendentalgroup.com', 'Veden Dental Group Inc');
$mail->AddAddress('m@vedendentalgroup.com', 'Veden Dental Group Inc');
// $mail->AddAddress('amanism@gmail.com', 'Veden Dental Group Inc');
$mail->AddAddress('it@vedendentalgroup.com', 'Veden Dental Group Inc');
*/
/*
//Set who the message is to be sent from
$mail->SetFrom('scan@vedendentalgroup.com', 'Veden Dental Group Inc');

//Set an alternative reply-to address
$mail->AddReplyTo('m@vedendentalgroup.com', 'Veden Dental Group Inc');

//Set who the message is to be sent to
$mail->AddAddress('amanism@gmail.com', 'Veden Dental Group Inc');

//Set the subject line
$mail->Subject = 'PHPMailer attachement test';

//Read an HTML message body from an external file, convert referenced images to embedded, convert HTML into a basic plain-text alternative body
// $mail->MsgHTML(file_get_contents('contents.html'), dirname(__FILE__));
//Replace the plain text body with one created manually
// $mail->AltBody = 'This is a plain-text message body';
$mail->Body = 'This is a plain-text message body -- FROM VEDENDENTALGROUP ';
//Attach an file
$mail->AddAttachment($filename);

if(!$mail->Send()) {
  echo "Mailer Error: " . $mail->ErrorInfo;
} else {
  echo "Message sent!";
}
*/
?>
