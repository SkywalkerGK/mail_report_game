<?php
/**
 * PHPExcel
 *
 * Copyright (c) 2006 - 2015 PHPExcel
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExcel
 * @package    PHPExcel
 * @copyright  Copyright (c) 2006 - 2015 PHPExcel (http://www.codeplex.com/PHPExcel)
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt	LGPL
 * @version    ##VERSION##, ##DATE##
 */

/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('Europe/London');

if (PHP_SAPI == 'cli')
	die('This example should only be run from a Web Browser');

/** Include PHPExcel */
require_once '/var/www/html/cacti_rep/Classes/PHPExcel.php';


// Create new PHPExcel object
$objPHPExcel = new PHPExcel();


//Connect MySQL
$con = mysqli_connect("xx.xx.xx.208","root","$-AJhxWP8@yyjv@6");

if (!$con) {
  die("Database connection failed: " . mysqli_error());
}
$db_select = mysqli_select_db($con, "icb_report");
if (!$db_select) {
  die("Database selection failed: " . mysqli_error());
}
mysqli_set_charset($con, 'utf8');

$year_now = date('Y');


$YESTERDATE = date("Y-m-01 00:00:00" , strtotime("-1 month")) ;
$DATE = date("Y-m-01 23:59:59");
$titledate = date('Y-m-d');

/* ---//////////////////////////////////////--- บรรทัดแรก Report Game : START ---//////////////////////////////////////---  */
  $row_current = 1; // บรรทัด
  $colum_next = 65; // รหัส Ascii A = 65
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'Report Game : '.$titledate); // Cell
  $colum_next += 1;

  /* ---//////////////////////////////////////--- MergeCells ---//////////////////////////////////////---  */
  $objPHPExcel->getActiveSheet()->mergeCells(chr(65).$row_current.':'.chr(73).$row_current); // การ mergeCells

    /* ------------------------------------------- ใส่ Set Text Center Bold : --------------------------------------------------------  */
    $StyleTextColorArray = array(
      'font'  => array(
          'bold'  => true,
          //'color' => array('rgb' => 'ffffff'), // ใส่ Code Color ที่จะใส่ให้ Text
          'size'  => 14,
          //'name'  => 'Verdana'
      ),
      'alignment' => array(
          'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
      )
    );

  $objPHPExcel->getActiveSheet()->getStyle(chr(65).$row_current.':'.chr(73).$row_current)->applyFromArray($StyleTextColorArray); // ใส่ Cell ที่จะเปลี่ยนสีตัวหนังสือ
  /* ------------------------------------------- ใส่ Set Text Center Bold : --------------------------------------------------------  */


/* ----------------------------------- ใส่ Set Cell Size   -----------------------------------  */
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(12);
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(5);
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('H')->setWidth(20);
$objPHPExcel->getActiveSheet()->getColumnDimension('I')->setWidth(70);
/* ----------------------------------- ใส่ Set Cell Size  -----------------------------------  */

/* ------------------------------------------------- Game Center -------------------------------------- */

$row_current += 2; // บรรทัด
$colum_next = 65; // รหัส Ascii A = 65
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'ticket_no'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'ro'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'provice'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'customer_name'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'dispatch_date'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'receieve_date'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'recieve_by'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'close_date'); // Cell
$colum_next++;
$objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, 'detail'); // Cell
$colum_next++;

$colum_next = 65; // รหัส Ascii A = 65

    /* ------------------------------------------- ใส่ Set Text Center Bold : --------------------------------------------------------  */
    $StyleTextColorArray = array(
      'font'  => array(
          'bold'  => true,
          //'color' => array('rgb' => 'ffffff'), // ใส่ Code Color ที่จะใส่ให้ Text
          //'size'  => 14,
          //'name'  => 'Verdana'
      )
    );
  $objPHPExcel->getActiveSheet()->getStyle(chr(65).$row_current.':'.chr(73).$row_current)->applyFromArray($StyleTextColorArray); // ใส่ Cell ที่จะเปลี่ยนสีตัวหนังสือ
  /* ------------------------------------------- ใส่ Set Text Center Bold : --------------------------------------------------------  */

$sql = mysqli_query($con, "select ticket_no,ro,provice,customer_name,dispatch_date, receieve_date, recieve_by,close_date,detail          
from fault_clearance_dis where remark like '%Game%' and receieve_date!='0000-00-00 00:00:00'          
and close_date between '$YESTERDATE' and '$DATE'");
$count = 0;

while ($result=mysqli_fetch_array($sql)) {
  $count++;
  $row_current += 1; // บรรทัด
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["ticket_no"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["ro"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["provice"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["customer_name"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["dispatch_date"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["receieve_date"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["recieve_by"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["close_date"]); // Cell
  $colum_next++;
  $objPHPExcel->getActiveSheet()->setCellValue(chr($colum_next).$row_current, $result["detail"]); // Cell

  $colum_next = 65; // รหัส Ascii A = 65
  //$text_month[] = $result["text_month"];
}

    /* ------------------------------------------- ใส่ Set Text Center Ro : --------------------------------------------------------  */
    $StyleTextColorArray = array(
      'alignment' => array(
          'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
      )
    );

  $objPHPExcel->getActiveSheet()->getStyle(chr(66).(4).':'.chr(66).$row_current)->applyFromArray($StyleTextColorArray); // ใส่ Cell ที่จะเปลี่ยนสีตัวหนังสือ
  /* ------------------------------------------- ใส่ Set Text Center Ro : --------------------------------------------------------  */

/* ------------------------------------------- ใส่ Borders Table  --------------------------------------------------------  */
$StyleBordersArray = array(
  'borders' => array(
    'allborders' => array(
      'style' => PHPExcel_Style_Border::BORDER_THIN
    )
  )
);
$objPHPExcel->getActiveSheet()->getStyle(chr(65).(3).':'.chr(73).$row_current)->applyFromArray($StyleBordersArray); // ใส่ borders
unset($StyleBordersArray); //ใส่ borders
/* ------------------------------------------- ใส่ Borders Table   --------------------------------------------------------  */


//////////////////////////////////////////////// END Table RADIUS ///////////////////////////////////////////////

// Redirect output to a client’s web browser (Excel5)
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment;filename="Summary_'.$year_now.''.'_ICB'.'.xls"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');

// If you're serving to IE over SSL, then the following may be needed
header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header ('Pragma: public'); // HTTP/1.0

$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
$objWriter->save('php://output');
exit;