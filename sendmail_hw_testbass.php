<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<?php

$dateall=date('Y-m-d');
//$strTo= "thodsaphon.k@jasmine.com,matis.a@jasmine.com,matis_3bb@yahoo.com,arttapol.r@jasmine.com,wanchai.ti@jasmine.com";
$strTo= "waree.l@jasmine.com";
#,songpon.j@jasmine.com

//$strSubject = "Test Report Game send excel file ภาษาไทย เดือน สิงหาคม";

$strSubject = "=?UTF-8?B?".base64_encode("Monthly Report Game ประจำวันที่ : ".$dateall)."?=";

$strMessage = "เรียน ทุกท่าน <br><br>
Monthly Report Game ประจำวันที่ : ".$dateall."<br><br><br>

<p>
<span style='color:#1F497D'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; *หมายเหตุ : </span>
<span style='color:red'>เป็นระบบส่งอีเมลอัตโนมัติ
    <br>
    <br>
</span>
</p>

<p>
<b>
    <span style='color:#002060'>ขอแสดงความนับถือ</span>
</b>
</p>

<p>

<b>
    <span style='color:gray'>o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o-o</span>
</b>
<span style='color:gray'>
    <o:p></o:p>
</span>
</p>


sorawit.sa@jasmine.com<br>
Tel . <br>
TripleT Broadband PCL.<br>
Jasmine International Tower, 21th Fl.<br>";

//*** Uniqid Session ***//
$strSid = md5(uniqid(time()));

$strHeader = "";
$strHeader .= "From: RADMIN<radmin@3bbmail.com>\nReply-To: radmin@3bbmail.com\n";
//$strHeader .= "Cc: sorawit.sa@jasmine.com\n";
$strHeader .= "Cc: radmin@3bbmail.com,pathumthip.p@jasmine.com,sorawit.sa@jasmine.com\n";

$strHeader .= "MIME-Version: 1.0\n";
$strHeader .= "Content-Type: multipart/mixed; boundary=\"".$strSid."\"\n\n";
$strHeader .= "This is a multi-part message in MIME format.\n";

$strHeader .= "--".$strSid."\n";
$strHeader .= "Content-type: text/html; charset=UTF-8\n"; // or UTF-8 //
$strHeader .= "Content-Transfer-Encoding: 7bit\n\n";
$strHeader .= $strMessage."\n\n";

//*** Attachment Files ***//
$arrFiles = 'report_game'.date("Ymd").'.xls';
if(trim($arrFiles) != "")
{
$strFilesName = $arrFiles;
$strContent = chunk_split(base64_encode(file_get_contents('/home/mapuser/report_game_bybass/csv/'.$strFilesName)));

$strHeader .= "--".$strSid."\n";
$strHeader .= "Content-Type: application/octet-stream; name=\"".$strFilesName."\"\n";
$strHeader .= "Content-Transfer-Encoding: base64\n";
$strHeader .= "Content-Disposition: attachment; filename=\"".$strFilesName."\"\n\n";
$strHeader .= $strContent."\n\n";
}

$flgSend = @mail($strTo,$strSubject,null,$strHeader); // @ = No Show Error //
if($flgSend)
{
echo "Email Sending.";
}
else
{
echo "Email Can Not Send.";
}

?>
