<?php

error_reporting(0);
$file_dir = "/files/filerepository/";

//DONT CHANGE ANYTHING BELOW HERE--------------------------------------------------------------------------------------------------------------------
header('Content-Type: text/html; charset= utf-8');
//$userDoc = $file_dir.$_POST['prefix'].$_POST['set1'].$_POST['set2'].$_POST['set3'].".doc";
$userDoc = $file_dir.$_POST['prefix'].$_POST['set1'].$_POST['set2'].$_POST['set3'];
//echo $userDoc;
//Captcha Check Start
require_once dirname(__FILE__) . '/securimage.php';
$securimage = new Securimage();

if ($securimage->check($_POST['ct_captcha']) == false) {
echo '<div class="success">
<div class="notice"><img src="/files/cert-images/error.png" width="64" height="64" align="absmiddle" />Incorrect security code entered</div>
</div>';
exit();
}
//Captcha Check End

//$userDoc = $file_dir."GB1067177509";
//$userDoc = $file_dir."GB1000100010";
//$userDoc = $file_dir."GB1067402612.doc";
//$userDoc = $file_dir."GB1067402512";



if (file_exists($userDoc.".doc") || file_exists($userDoc.".xls")) {

if(file_exists($userDoc.".doc")){
$text = parseWord($userDoc.".doc");//Method 2
//$html = nl2br(htmlspecialchars($text));
//$html = preg_replace('/\s\s+/', ' ', $html);
echo '<div class="success">
      <div class="notice"><img src="/files/cert-images/success.png" width="64" height="64" align="absmiddle" /><strong>Success!</strong> Certificate Data Found</div>
    </div>
	<p>'.$text.'</p></br><div class="footer">Certificate of adequacy is defined in <a href="http://www.legislation.gov.uk/uksi/1992/3073/regulation/20/made" target="_new">\'The Supply of Machinery (Safety) Regulations 1992\'</a></div><p></br></p>';
}
else{
$text = parseExcel($userDoc.".xls");
echo '<style>
table.excel {
	border-style:ridge;
	border-width: 0px ;
	border-collapse:collapse;
	font-family:sans-serif;
	font-size:14px !important;
	background: ;
	margin-left: auto;
    margin-right: auto;
    text-align: left;
    width: 800px;
}
table.excel thead th, table.excel tbody th {
	background: transparent;
	border-style:ridge;
	border-width: 0px ;
	text-align: center;
	vertical-align:bottom;

}
table.excel tbody th {
	text-align:center;
	width:20px;
	color: #000000;
}
table.excel tbody td {
	vertical-align:bottom;
	color: #000000;
}
table.excel tbody td {
    padding: 0 3px;
	border: 0px solid #EEEEEE;
	color: #000000;
}
.excel tbody tr td nobr {
	width: 0px;
	height: 10px;
	letter-spacing: 0px;
	font-size: 10px;
	color: #FFF;
}

</style>
<div class="success">
      <div class="notice"><img src="/files/cert-images/success.png" width="64" height="64" align="absmiddle" /><strong>Success!</strong> Certificate Data Found</div>
    </div>
	<p>'.$text.'</p></br><div class="footer">Certificate of adequacy is defined in <a href="http://www.legislation.gov.uk/uksi/1992/3073/regulation/20/made" target="_new">\'The Supply of Machinery (Safety) Regulations 1992\'</a></div><p></br></p>';
}







}
else{
echo '   <div class="success">
      <div class="notice"><img src="/files/cert-images/error.png" width="64" height="64" align="absmiddle" /><strong>Sorry</strong> We couldn\'t find that Certificate</div>
    </div>
    <p><div class="footer">Please check the Certificate number and try again, or please <a href="http://www.avtechnology.co.uk/contacts.php">contact us</a> for a manual validation.</div><p></br></p>
   <p>&nbsp;</p>';

}


/*****************************************************************
This approach uses detection of NUL (chr(00)) and end line (chr(13))
to decide where the text is:
- divide the file contents up by chr(13)
- reject any slices containing a NUL
- stitch the rest together again
- clean up with a regular expression
*****************************************************************/





function parseWord($userDoc)
{
    $fileHandle = fopen($userDoc, "r");
    $word_text = @fread($fileHandle, filesize($userDoc));
    $line = ""; $lineord = "";
    $tam = filesize($userDoc);
    $nulos = 0;
    $caracteres = 0;
   // for($i=1536; $i<$tam; $i++)
    for($i=2536; $i<$tam; $i++)
    {
        $line .= $word_text[$i];

		if( ord($word_text[$i]) != 0)
		$lineord .= ord($word_text[$i]);

        if( $word_text[$i] == 0)
        {
            $nulos++;
        }
        else
        {
            $nulos=0;
            $caracteres++;
        }

        if( $nulos>1996 || strstr($lineord,"13131313"))
        {
            break;
        }
    }

    //echo $caracteres;

    $lines = explode(chr(0x0D),$line);
    //$outtext = "<pre>";

    $outtext = "";
    foreach($lines as $thisline)
    {
        $tam = strlen($thisline);
        if( !$tam )
        {
            continue;
        }

        $new_line = "";
        for($i=0; $i<$tam; $i++)
        {
            $onechar = $thisline[$i];
            if( $onechar > chr(240) )
            {
                continue;
            }

            if( $onechar >= chr(0x20) )
            {
                $caracteres++;
                $new_line .= $onechar;
            }

            if( $onechar == chr(0x14) )
            {
                $new_line .= "</a>";
            }

            if( $onechar == chr(0x07) )
            {
                //$new_line .= "\t";
				$new_line .= "";
                if( isset($thisline[$i+1]) )
                {
                    if( $thisline[$i+1] == chr(0x07) )
                    {
                        $new_line .= "\n";
						//$new_line .= "<br>";
                    }
                }
            }
        }
        //troca por hiperlink
        $new_line = str_replace("HYPERLINK" ,"<a href=",$new_line);
        $new_line = str_replace("\o" ,">",$new_line);
        $new_line .= "\n";

        //link de imagens
        $new_line = str_replace("INCLUDEPICTURE" ,"<br><img src=",$new_line);
        $new_line = str_replace("\*" ,"><br>",$new_line);


		//$new_line = str_replace("&nbsp;" ," ",$new_line);
        $new_line = str_replace("MERGEFORMATINET" ,"",$new_line);
		//$new_line=preg_replace('`[\r\n]+`',"\n",$new_line);

		//$new_line = preg_replace('/\n\s+/', "\n", $new_line );
		//$new_line =preg_replace('#<br>(\s*<br>)+#', '<br>', $new_line);
		//$new_line = preg_replace('~(^<br>\s*)|((?<=<br>)\s*<br>)|(<br>\s*$)~', '', $new_line);
		//


		$outtext .= nl2br($new_line);



 // nl2br($new_line);



    }
$outtext = str_replace("\r","",$outtext);
$outtext = str_replace("\n","",$outtext);

$outtext = preg_replace('/(<br\s*\/?>\s*)+/', "<br>", $outtext);

 return $outtext;
}


function parseExcel($userDoc)
{
require_once 'excel_reader2.php';
$data1 = new Spreadsheet_Excel_Reader($userDoc);
return $data1->dump($row_numbers=false,$col_letters=false);
}

?>
