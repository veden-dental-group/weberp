<?php 

/** 
* class Mailer 
* 
* Class to generate a mail with html AND plain text message. 
* Also possible to attach files 
* 
* Methods: 
* --------------------------------------------------------------------- 
* mailer()                                  - Constructor, set some default values 
* setHTMLMessage( $message )                - Set the HTML message 
* setMessage( $message )                    - Set a plain text message 
* attachFile( $file [, $mime] )             - Attach a file to the mail 
* includeFile( $file [, $mime] )            - Attach a file which can be used in the HTML message. The name which can be used in the HTML will be returned 
* attachFileAsData( $file, $data, $mime )   - Attach a file and give the contents of the file as data. 
* includeFileAsData( $file, $data, $mime )  - Attach a file and give the contents of the file as data. The file can be used in the HTML message. The name which can be used in the HTML will be returned 
* setParameters( $paramenters )             - Set extra parameters for the mail function 
* setBoundary( $boundary )                  - Set a specific boundary 
* setPriority( $priority )                  - Set the priority of the message 
* setFrom( $mail [, $name] )                - Set the from e-mail adres and name 
* setReplyTo( $mail [, $name] )             - Set the reply-to e-mail adres and name 
* setCc( $cc )                              - Set the cc addresses to send a copy of a mail to 
* setBcc( $bcc )                            - Set the bcc addresses to send a "blind" copy of the mail 
* setHeader( $header )                      - Set extra headers to the mail 
* sendReadConfirmationTo( $mail [, $name] ) - Send a read confirmation to this address 
* send( $to, $subject )                     - Send the mail 
* 
* Example: 
* ---------------------------------------------------------------------- 
* $m = new mailer(); 
* 
* $img = $m->includeFile( "logo.gif", "image/gif" ); 
* 
* $m->setHTMLMessage( "This is a <b>test</b> e-mail <i>with</i> HTML! <img src=\"$img\" />" ); 
* $m->setMessage("This is a plain text message..."); 
* 
* $csvData = 'test;test2;someotherData;'; 
* 
* $m->attachFileAsData("myCSVFile", $csvData, "text/plain"); 
* 
* $m->setPriority( 'High' ); 
* 
* $m->setFrom( "john@company.com", "J. Doe" ); 
* $m->setReplyTo( "info@company.com", "My Company Name" ); 
* 
* if( $m->send("john@doe.com", "Test Mail!" ) ) { 
*     echo "Mail is send!"; 
* } else { 
*     echo "Something went wrong! Please try again!"; 
* } 
* 
* @author Teye Heimans 
* @package FormHandler 
*/ 
class Mailer 
{ 
    var $_headers; 
    var $_boundary; 
    var $_message; 
    var $_params; 
    var $_fileTypes; 

    /** 
     * mailer::mailer() 
     * 
     * Constructor, set some default values. 
     * 
     * @access protected 
     * @author Teye Heimans 
     */ 
    function mailer() 
    { 
        // always use \r\n in your mail and to split the headers! 
        if( !defined('NL') ) { 
            define('NL', "\n"); 
        } 

        // create a unique boundary 
        $this->setBoundary( 'myMailBoundary('.uniqid(time()).')' ); 

        // Set the default text message 
        // this will be overwritten if the user set's another 
        // plain text message 
        $this->setMessage( 
          'If you are reading this, it means that your e-mail client'.NL. 
          'does not support MIME. Please upgrade your e-mail client'.NL. 
          'to one that supports MIME.' 
        ); 

        // set the from adres when it can be retreived from the php.ini 
        $from = ini_get('sendmail_from'); 
        if( !empty($from) ) 
        { 
            $this->setFrom( $from ); 
        } 

        // set some basic headerdata 
        $this->setHeader( 
          'X-Mailer: PHP '.phpversion() . NL . 
          'X-MimeOLE: PHP '.phpversion() . NL . 
          'Date: '.date('r') . NL . 
          'Message-ID: <'.md5(uniqid(time())).'@'.$_SERVER['SERVER_NAME'].'>' 
        ); 

        // set the priority to normal 
        $this->setPriority( 3 ); 

        // this types will be used for auto loading the mime type 
        $this->image_types = array( 
          'gif'    => 'image/gif', 
          'jpg'    => 'image/jpeg', 
          'jpeg'    => 'image/jpeg', 
          'jpe'    => 'image/jpeg', 
          'bmp'    => 'image/bmp', 
          'png'    => 'image/png', 
          'tif'    => 'image/tiff', 
          'tiff'    => 'image/tiff', 
          'swf'    => 'application/x-shockwave-flash' 
        ); 

    } 

    /** 
     * mailer::setHTMLMessage() 
     * 
     * Set the HTML message. Most clients do not accept HTML e-mails, so 
     * you should ALWAYS set a plain text message! 
     * 
     * @param string $message: The HTML message to set 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setHTMLMessage( $message ) 
    { 
        // check message for mail injection 
        if( !preg_match('/Content-Type:/i', $message) ) 
        { 
            $this->_message['html'] = $message; 
        } 
    } 

    /** 
     * mailer::setMessage() 
     * 
     * Set the plain text message. Most clients do not accept HTML e-mails, so 
     * you should ALWAYS set a plain text message! 
     * 
     * @param string $message: The message to set 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setMessage( $message ) 
    { 
        // check message for mail injection 
        if( !preg_match('/Content-Type:/i',$message) ) 
        { 
            $this->_message['plain'] = $message; 
        } 
    } 


    /** 
     * mailer::attachFile() 
     * 
     * Attach a file to the mail 
     * 
     * @param string $file: the file we have to attach to the mail 
     * @param string $mime: if the mime type cant be retrieved you have to spicify the mime type of the file 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function attachFile( $file, $mime = null ) 
    { 
        // retieve the mime type of the file if possible 
        if( function_exists('mime_content_type') ) 
        { 
            $mime = mime_content_type( $file ); 

        } 
        // if the mime type can not be retrieved and it is not given, 
        // display an error message 
        if ( is_null( $mime) || !$mime ) 
        { 
            trigger_error( 
              'The mime type of the file "'.$file.'" could not be retrieved.' . NL . 
              'Please set the mime type of the file as second argument in the method attachFile().' . NL . 
              'If you don\'t know the mime type, you can retireve it at http://filext.com.', 
              E_USER_WARNING 
            ); 
            return; 
        } 

        // check if the file exists and if it is readable and open it 
        if( !file_exists( $file ) || !is_readable( $file ) || !($fd = fopen($file, 'rb')) ) 
        { 
            trigger_error( 
              'The file "'.$file.'" could not be read. Make sure that the '. 
              'file exists and that is readable.', 
              E_USER_WARNING 
            ); 
            return; 
        } 

        // read the data from the file 
        $data = fread( $fd, filesize($file) ); 
        fclose( $fd ); 

        // get the name of the file 
        $name = basename($file); 

        // create the attachment data 
        $this->_message['attachment'][] = 
          'Content-Type: '.$mime.'; name="'.$name.'"' . NL . 
          'Content-Transfer-Encoding: base64' . NL . 
          'Content-disposition: attachment; file="'.$name.'"' . NL . NL . 
          chunk_split( base64_encode( $data ) ) . NL ; 
    } 

    /** 
     * mailer::includeFile() 
     * 
     * HTML emails will attempt to download their images and style sheets from the web. 
     * Because of security and privacy reasons, many clients will refuse to attempt these downloads, 
     * ruining the look of your HTML message. With this function you can attach the files 
     * to your e-mail and point to them in your HTML. 
     * 
     * @param string $file: the file we have to attach to the mail 
     * @param string $mime: if the mime type cant be retrieved you have to spicify the mime type of the file 
     * @return string: the name you can use in your HTML tag, like <img src="xxx" /> 
     * @access public 
     * @author Teye Heimans 
     */ 
    function includeFile( $file, $mime ) 
    { 
        // retieve the mime type of the file if possible 
        if( function_exists('mime_content_type') ) 
        { 
            $mime = mime_content_type( $file ); 

        } 
        // if the mime type can not be retrieved and it is not given, 
        // display an error message 
        if ( is_null( $mime) || !$mime ) 
        { 
            trigger_error( 
              'The mime type of the file "'.$file.'" could not be retrieved.' . NL . 
              'Please set the mime type of the file as second argument in the method attachFile().' . NL . 
              'If you don\'t know the mime type, you can retireve it at http://filext.com.', 
              E_USER_WARNING 
            ); 
            return; 
        } 

        // check if the file exists and if it is readable and open it 
        if( !file_exists( $file ) || !is_readable( $file ) || !($fd = fopen($file, 'rb')) ) 
        { 
            trigger_error( 
              'The file "'.$file.'" could not be read. Make sure that the '. 
              'file exists and that is readable.', 
              E_USER_WARNING 
            ); 
            return; 
        } 


        // read the data from the file 
        $data = fread( $fd, filesize($file) ); 
        fclose( $fd ); 

        // get the name of the file as we are going to use it in the html email 
        $name = in_array(substr($file, 0, 4), array('http', 'ftp:')) ? $file : uniqid('') .'/'. basename( $file ); 

        // create the attachment data 
        $this->_message['include'][] = 
          'Content-Type: '.$mime.'; name="'.$name.'"' . NL . 
          'Content-Transfer-Encoding: base64' . NL . 
          'Content-Location: '.$name . NL . NL . 
          chunk_split( base64_encode( $data ) ) . NL ; 

        return $name; 
    } 


    /** 
     * mailer::setPriority() 
     * 
     * Set the priority of the e-mail (Default normal). 
     * You can use numbers (1 = highest, 5 = lowest) or textual strings: ("highest" - "lowest") 
     * 
     * @param mixed $priority: The priotiry for the mail 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setPriority( $priority = 3 ) 
    { 
        $priorities = array( 
          '1 (Highest)', 
          '2 (High)', 
          '3 (Normal)', 
          '4 (Low)', 
          '5 (Lowest)' 
        ); 

        switch( strtolower($priority) ) 
        { 
            case 'highest': 
            case 1: 
              $a = $priorities[0]; 
              $b = 'Highest'; 
              break; 

            case 'high': 
            case 2: 
              $a = $priorities[1]; 
              $b = 'High'; 
              break; 

            case 'low': 
            case 4: 
              $a = $priorities[3]; 
              $b = 'Low'; 
              break; 

            case 'lowest': 
            case 5: 
              $a = $priorities[4]; 
              $b = 'Lowest'; 
              break; 

            case 3: 
            case 'normal': 
            default: 
              $a = $priorities[2]; 
              $b = 'Normal'; 
              break; 
        } 

        $this->setHeader( 
          'X-Priority: ' . $a . NL. 
          'X-MSMail-Priority: ' . $b . NL 
        ); 
    } 

    /** 
     * mailer::attachFileAsData() 
     * 
     * Attach a file by giving the binary data to this method. 
     * (The user should fetch the file data and the mime type) 
     * 
     * @param string $file: The name of the file we should attach 
     * @param string $data: The binary data of the file which we should attach 
     * @param string $mime: the mime type of the file you want to attach 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function attachFileAsData( $file, $data, $mime ) 
    { 
        // get the name of the file 
        $name = basename( $file ); 

        // create the attachment data 
        $this->_message['attachment'][] = 
          'Content-Type: '.$mime.'; name="'.$name.'"' . NL . 
          'Content-Transfer-Encoding: base64' . NL . 
          'Content-disposition: attachment; file="'.$name.'"' . NL . NL . 
          chunk_split( base64_encode( $data ) ) . NL ; 
    } 


    /** 
     * mailer::includeFileAsData() 
     * 
     * Includes a file by giving the binary data to this method. 
     * (The user should fetch the file data and the mime type) 
     * 
     * @param string $file: The name of the file we should include 
     * @param string $data: The binary data of the file which we should include 
     * @param string $mime: the mime type of the file you want to include 
     * @return string: The name which can be used in the HTML message 
     * @access public 
     * @author Teye Heimans 
     */ 
    function includeFileAsData( $file, $data, $mime ) 
    { 
        // get the name of the file as we are going to use it in the html email 
        $name = in_array(substr($file, 0, 4), array('http', 'ftp:')) ? $file : uniqid('') .'/'. basename( $file ); 

        // create the attachment data 
        $this->_message['include'][] = 
          'Content-Type: '.$mime.'; name="'.$name.'"' . NL . 
          'Content-Transfer-Encoding: base64' . NL . 
          'Content-Location: '.$name . NL . NL . 
          chunk_split( base64_encode( $data ) ) . NL ; 

        return $name; 

    } 

    /** 
     * mailer::setParameters() 
     * 
     * Set additional parameters for the mail function 
     * 
     * @param string $params: the additional params 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setParameters( $params ) 
    { 
        $this->_params = $params; 
    } 

    /** 
     * mailer::send() 
     * 
     * Create the mail and send it 
     * 
     * @param $to: The e-mail adres where we should send the mail to 
     * @param $subject: The subject of the mail 
     * @return bool 
     * @access public 
     * @author Teye Heimans 
     */ 
    function send( $to, $subject ) 
    { 
        // is the mail a mime mail? 
        $mime = false; 

        // does the mail contains included attachemnts? 
        $includes = false; 

        // set the plain text body 
        $body = $this->_message['plain'] . NL; 

        // is there a html message ? 
        if( isset( $this->_message['html'] ) ) 
        { 
            // the mail is a mime mail.. 
            $mime = true; 

            // set the html body 
            $body .= 
            '--' . $this->_boundary . NL . 
            'Content-Type: text/html; charset="iso-8859-1"'. NL . 
            'Content-Transfer-Encoding: quoted-printable'. NL . NL . 
            $this -> quoted_printable_encode( $this -> _message['html'] ) . NL; 


            /*'Content-Type: text/html; charset="iso-8859-1"' . NL . 
            'Content-Transfer-Encoding: base64' . NL . NL . 
            chunk_split(base64_encode($this->_message['html'])) . NL; 
            */ 
        } 

        // are there attachments ? 
        if( isset( $this->_message['attachment'] ) ) 
        { 
            // the mail is a mime mail.. 
            $mime = true; 

            foreach( $this->_message['attachment'] as $data ) 
            { 
                $body .= '--' . $this->_boundary . NL . $data; 
            } 
        } 

        // are there attachments ? 
        if( isset( $this->_message['include'] ) ) 
        { 
            // the mail is a mime mail.. 
            $mime = true; 

            // we use includes (embedded images) in our mail.. 
            $includes = true; 

            foreach( $this->_message['include'] as $data ) 
            { 
                $body .= '--' . $this->_boundary . NL . $data; 
            } 

        } 

        // if it's a mime mail, set the end boundary 
        if( $mime ) 
        { 
            $body .= '--' . $this->_boundary . '--' . NL; 

            // set the mime headers if it's a mime mail 
            $this->setHeader( 
              'MIME-Version: 1.0' . NL . 
              'Content-Type: multipart/'.($includes ? 'related':'mixed').'; boundary="'.$this->_boundary.'"' . NL 
            ); 
        } 

        if( !isset( $this->_params ) ) 
        { 
            $this->_params = null; 
        } 

        /** 
         * send the mail! 
         */ 

        // params set ? 
        if( isset( $this -> _params ) ) 
        { 
            // send the mail with the extra params! 
            $mail =  mail( 
              $to, 
              $subject, 
              $body, 
              $this->_getHeaders(), 
              $this -> _params 
            ); 
        } 
        // no params given.. just send the mail! 
        else 
        { 
            $mail =  mail( 
              $to, 
              $subject, 
              $body, 
              $this->_getHeaders() 
            ); 
        } 

        // put the original value back 
        if( isset( $this->_sendmail_from ) ) 
        { 
            ini_set('sendmail_from', $this->_sendmail_from ); 
        } 

        return $mail; 
    } 

    /** 
     * mailer::setBoundary() 
     * 
     * Set the boundary which is used to split the message in blocks. The 
     * boundary has to be a unique string which will not occur in the messages! 
     * Only use this function if you know what you are doing. 
     * The boundary which is set by default works fine. 
     * 
     * @param string $boundary: The new boundary 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setBoundary( $boundary ) 
    { 
        $this->_boundary = $boundary; 
    } 


    /** 
     * mailer::setFrom() 
     * 
     * Set the "From" adres of the mail 
     * 
     * @param string $email: The e-mail adres of the sender 
     * @param string $name: The name of the sender 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setFrom( $email, $name = '' ) 
    { 
        if( empty( $name ) ) { 
            $name = $email; 
        } 

        // make sure that the data is secure (for e-mail injection) 
        if( $this->_secure( $name, $email ) ) 
        { 
            $this->setHeader('From: '. $name.' <'.$email.'>' . NL ); 
            $this->setHeader('X-Sender: '.$email . NL ); 
            $this->setHeader('Return-Path: '.$email . NL ); 

            // save the original sendmail_from 
            $this->_sendmail_from = ini_get('sendmail_from'); 

            // set the new from 
            ini_set('sendmail_from', $email); 
        } 
    } 

    /** 
     * mailer::sendReadConfirmationTo() 
     * 
     * Send a read confirmation to the given adres 
     * 
     * @param string $mail: The e-mail address where the read confirmation should send to 
     * @param string $name: The name of the receiver 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function sendReadConfirmationTo( $mail, $name = '' ) 
    { 
        if( empty( $name ) ) { 
            $name = $email; 
        } 

        // make sure that the data is secure (for e-mail injection) 
        if( $this->_secure( $mail, $name) ) 
        { 
            $this->setHeader( 
              'Disposition-Notification-To: '.$name.' <'.$email.'>' 
            ); 
        } 
    } 

    /** 
     * mailer::setReplyTo() 
     * 
     * Set the "reply-to" adres for the mail 
     * 
     * @param string $email: The e-mail adres where the receiver should reply to 
     * @param string $name: The name of the sender 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setReplyTo( $email, $name = '' ) 
    { 
        if( empty( $name ) ) { 
            $name = $email; 
        } 

        // make sure that the data is secure (for e-mail injection) 
        if( $this->_secure( $name, $email ) ) 
        { 
            $this->setHeader('Reply-To: '.$name.' <'.$email.'>'.NL ); 
        } 
    } 

    /** 
     * mailer::setCc() 
     * 
     * Send a "carbon copy" to the given email addresses. 
     * Split multiple e-mail addresses with a space or a comma. 
     * (CC = a copy of the mail, the receiver can see to who the mail is also send) 
     * 
     * @param string $cc: The e-mail address(es) of the persons who should recieve a copy of the mail 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setCc( $cc ) 
    { 
        if( $this->_secure( $cc ) ) 
        { 
            $this->setHeader( 'Cc: ' . $cc . NL ); 
        } 
    } 

    /** 
     * mailer::setBcc() 
     * 
     * Send a "blind carbon copy" to the given email addresses. 
     * Split multiple e-mail addresses with a space or a comma. 
     * (CC = a copy of the mail, the receiver can NOT see to who the mail is also send) 
     * 
     * @param string $bcc: The e-mail address(es) of the persons who should recieve a copy of the mail 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setBcc( $bcc ) 
    { 
        if( phpversion() < 4.3 && strpos( 'win', strtolower($_SERVER['SERVER_SOFTWARE'])) !== false ) 
        { 
            trigger_error( 
              'Error, could not set the bcc addresses. The PHP version does not support this!', 
              E_USER_WARNING 
            ); 
            return; 
        } 

        if( $this->_secure( $bcc ) ) 
        { 
            $this->setHeader( 'Bcc: ' . $bcc . NL ); 
        } 
    } 

    /** 
     * mailer::setHeader() 
     * 
     * Add an extra header to the mail 
     * 
     * @param string $header: The extra header you want to add 
     * @return void 
     * @access public 
     * @author Teye Heimans 
     */ 
    function setHeader( $header ) 
    { 
        // split all headers 
        $headers = split("\r?\n", $header); 
        foreach ( $headers as $header ) 
        { 
            // save the headers in the header array 
            if( !empty( $header ) ) 
            { 
                list( $header, $data ) = explode( ':', $header, 2); 
                $header = trim($header); 

                $this->_headers[$header] = $data . NL; 
            } 
        } 
    } 

    /** 
     * mailer::_getHeaders() 
     * 
     * Return the headers in the correct format for the mail function 
     * 
     * @return string 
     * @access private 
     * @author Teye Heimans 
     */ 
    function _getHeaders( ) 
    { 
        // the result string 
        $result = ''; 

        // walk each header 
        foreach( $this->_headers as $name => $data ) 
        { 
            // upper case all first characters (also after an "-" char) 
            foreach ( explode('-', $name) as $str) 
            { 
                $result .= ucfirst( $str ) . '-'; 
            } 

            // make sure that the line ends with an enter 
            if( !preg_match( "/\r?\n$/", $data) ) 
            { 
                $data .= NL; 
            } 

            // add the header to the result 
            $result = substr($result, 0, -1) . ': ' . $data; 
        } 

        // return the result (last header has to have 2 NL's! 
        return $result . NL ; 
    } 

    /** 
     * mailer::_secure() 
     * 
     * Check if the given arguments are secure (it may not contain \r or \n 
     * because of e-mail injection and not "MIME-Version" or "Concent-Type") 
     * 
     * @return boolean 
     * @access private 
     * @author Teye Heimans 
     */ 
    function _secure() 
    { 
        // number of arguments 
        $num = func_num_args(); 

        // walk all arguments 
        for( $i = 0; $i < $num; $i++ ) 
        { 
            $arg = func_get_arg( $i ); 
            if( 
              preg_match("/\r|\n/", $arg ) || 
              preg_match('/MIME-Version:/i', $arg ) || 
              preg_match('/Content-Type:/i', $arg)) 
            { 
                return false; 
            } 
        } 

        return true; 
    } 

    /** 
     * Mailer::quoted_printable_encode() 
     * 
     * Return a given string so that it can be used 
     * with a Content-Transfer-Encoding set to quoted-printable. 
     * 
     * @param string $input: the text which should be converted to quoted-string 
     * @param int $line_max: the number of characters allowed on a line 
     * @param bool $space_conv: Should a space be converted also ? 
     * @return string 
     * @access private 
     */ 
    function quoted_printable_encode( $sText, $bEmulate_imap_8bit = true ) 
    { 
        // split text into lines 
        $aLines = explode( chr(13).chr(10), $sText ); 

        // walk al lines 
        for( $i = 0; $i < count($aLines); $i++) 
        { 
            $sLine =& $aLines[$i]; 
            if (strlen($sLine)===0) continue; // do nothing, if empty 

            $sRegExp = '/[^\x09\x20\x21-\x3C\x3E-\x7E]/e'; 

            // imap_8bit encodes x09 everywhere, not only at lineends, 
            // for EBCDIC safeness encode !"#$@[\]^`{|}~, 
            // for complete safeness encode every character :) 
            if ($bEmulate_imap_8bit) 
            { 
                $sRegExp = '/[^\x20\x21-\x3C\x3E-\x7E]/e'; 
            } 


            $sReplmt = 'sprintf( "=%02X", ord ( "$0" ) ) ;'; 
            $sLine = preg_replace( $sRegExp, $sReplmt, $sLine ); 

            // encode x09,x20 at lineends 
            $iLength = strlen($sLine); 
            $iLastChar = ord($sLine{$iLength-1}); 

            //              !!!!!!!! 
            // imap_8_bit does not encode x20 at the very end of a text, 
            // here is, where I don't agree with imap_8_bit, 
            // please correct me, if I'm wrong, 
            // or comment next line for RFC2045 conformance, if you like 
            if (!($bEmulate_imap_8bit && ($i==count($aLines)-1))) 
            { 
                if (($iLastChar==0x09)||($iLastChar==0x20)) 
                { 
                    $sLine{$iLength-1}='='; 
                    $sLine .= ($iLastChar==0x09)?'09':'20'; 
                } 
            } 

            // imap_8bit encodes x20 before chr(13), too 
            // although IMHO not requested by RFC2045, why not do it safer :) 
            // and why not encode any x20 around chr(10) or chr(13) 
            if ($bEmulate_imap_8bit) 
            { 
                $sLine=str_replace(' =0D','=20=0D',$sLine); 
                //$sLine=str_replace(' =0A','=20=0A',$sLine); 
                //$sLine=str_replace('=0D ','=0D=20',$sLine); 
                //$sLine=str_replace('=0A ','=0A=20',$sLine); 
            } 

            // finally split into softlines no longer than 76 chars, 
            // for even more safeness one could encode x09,x20 
            // at the very first character of the line 
            // and after soft linebreaks, as well, 
            // but this wouldn't be caught by such an easy RegExp 
            preg_match_all( '/.{1,73}([^=]{0,2})?/', $sLine, $aMatch ); 
            $sLine = implode( '=' . chr(13).chr(10), $aMatch[0] ); // add soft crlf's 
        } 

        // join lines into text 
        return implode(chr(13).chr(10),$aLines); 
    } 
} 
?> 
