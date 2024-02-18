<?php
//Import PHPMailer classes into the global namespace
//These must be at the top of your script, not inside a function
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\IOFactory;

//Load Composer's autoloader
require 'vendor/autoload.php';

//Initialization of paths
$excel_path = "Excel/tabel.xlsx";

//load excel table
$spreadsheet = IOFactory::load($excel_path);
$sheet = $spreadsheet->getActiveSheet();

//Create an instance; passing `true` enables exceptions
$mail = new PHPMailer(true);

//Server settings
$mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
$mail->isSMTP();                                            //Send using SMTP
$mail->Host       = 'mail.stepbystepeducation.ro';                     //Set the SMTP server to send through
$mail->SMTPAuth   = true;                                   //Enable SMTP authentication
$mail->Username   = 'congres@stepbystepeducation.ro';                     //SMTP username
$mail->Password   = 'B0QpJfNXwk';                               //SMTP password
$mail->SMTPSecure = PHPMailer::ENCRYPTION_SMTPS;            //Enable implicit TLS encryption
$mail->Port       = 465;

//Settings
$mail->setFrom('congres@stepbystepeducation.ro', 'StepbyStepEducation');
$mail->addReplyTo('congres@stepbystepeducation.ro', 'Replytome');

//Subject and body
$mail->isHTML(true);                                  //Set email format to HTML
$mail->Subject = 'Congres StepByStep Februarie 2024 - Certificat de Participare';
$mail->Body    = 'Buna ziua, am emis si atasat in email Certificatul de Participare.';

// Get name and email column
$nameColumn = 'A';
$emailColumn = 'D';
$checkColumn = 'R';
$highestRow = $sheet->getHighestRow();

for ($row = 2; $row <= 398; $row++) {+
    $emailValue= $sheet->getCell($emailColumn . $row)->getValue();
    $nameValue = $sheet->getCell($nameColumn . $row)->getValue();
    if (!empty($nameValue)) {
        //replace diacritics
        $nameValue = str_replace('Ă', 'a', $nameValue);
        $nameValue = str_replace('Â', 'a', $nameValue);
        $nameValue = str_replace('Î', 'i', $nameValue);
        $nameValue = str_replace('Ș', 's', $nameValue);
        $nameValue = str_replace('Ț', 't', $nameValue);
        $nameValue = str_replace('ă', 'a', $nameValue);
        $nameValue = str_replace('â', 'a', $nameValue);
        $nameValue = str_replace('î', 'i', $nameValue);
        $nameValue = str_replace('ș', 's', $nameValue);
        $nameValue = str_replace('ț', 't', $nameValue);

        //get the recipient email from excel if it exists and is valid
        try {
            $mail->addAddress($emailValue, $nameValue);
        } catch (Exception $e) {
            echo 'Invalid address skipped: ' . htmlspecialchars($emailValue) . '<br>';
            continue;
        }
        //add the attachement to be sent
        $mail->addAttachment('Diplome/' . $emailValue . ".pdf" , 'Diploma.pdf');

        //try sending the email
        try {
            $mail->send();
            echo 'Message sent to :' . htmlspecialchars($nameValue) . ' (' .
                htmlspecialchars($emailValue) . ')<br>';
                //Mark it as sent in the excel
                $sheet->getCell($checkColumn . $row)->setValue('Sent');
        } catch (Exception $e) {
            //if the email is not sent mark it as such in the excel and the console
            echo 'Mailer Error (' . htmlspecialchars($emailValue) . ') ' . $mail->ErrorInfo . '<br>';
            $sheet->getCell($checkColumn . $row) -> setValue('Failed');
            //Reset the connection to abort sending this message
            //The loop will continue trying to send to the rest of the list
            $mail->getSMTPInstance()->reset();
        }
        //clear everything and save the excel
        $mail->clearAddresses();
        $mail->clearAttachments();
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('Excel/tabel.xlsx');

    } else {
        //save the excel and break if there are no more values in the table
        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('Excel/tabel.xlsx');
        break;
    }
}
//ssave the excel again just to be safe
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('Excel/tabel.xlsx');

exit();
