<?php
require_once '../excel_reader2.php';
require_once("../PHPExcel.php");

date_default_timezone_set('Asia/Ho_Chi_Minh');

if( $_FILES['file']['tmp_name']) 
{
    global $dulieu;
    global $ketqua;
    global $export;
    $sodoi = $_POST['number'];
    $data = new Spreadsheet_Excel_Reader($_FILES['file']['tmp_name']);
    $rowsnum = $data->rowcount($sheet_index=0); 
    $colsnum =  $data->colcount($sheet_index=0);
    for ($i = 2; $i <= $rowsnum; $i++)
    {
        $dulieu[$i]['tenhang'] = $data->val($i,1,0);
        $dulieu[$i]['mahang'] = $data->val($i,2,0);
        $dulieu[$i]['size'] = $data->val($i,3,0);
        $dulieu[$i]['soluong'] = $data->val($i,4,0);
        $dulieu[$i]['thungdu'] = floor($data->val($i,4,0)/$sodoi);
        $dulieu[$i]['thungthieu'] = $data->val($i,4,0)%$sodoi;
    }
    $so = 1;
    $so2 = 1;
    $sothutu = 1;
    foreach($dulieu as $value){
        for($i = 0; $i < $value['thungdu']; $i++){
            $ketqua[$so][$so2]['tenhang'] = $value['tenhang'];
            $ketqua[$so][$so2]['mahang'] = $value['mahang'];
            $ketqua[$so][$so2]['size'] = $value['size'];
            $ketqua[$so][$so2]['soluong'] = $sodoi;
            $ketqua[$so][$so2]['sothutu'] = $sothutu;
            $so2++;
            $sothutu++;
            if($so2 > 6){
                $so++;
                $so2=1;
            }
        }
        if($value['thungthieu'] > 0){
        $ketqua[$so][$so2]['tenhang'] = $value['tenhang'];
        $ketqua[$so][$so2]['mahang'] = $value['mahang'];
        $ketqua[$so][$so2]['size'] = $value['size'];
        $ketqua[$so][$so2]['soluong'] = $value['thungthieu'];
        $ketqua[$so][$so2]['sothutu'] = $sothutu;
        $so2++;
        $sothutu++;
        if($so2 > 6){
                $so++;
                $so2=1;
            }
        }
        $sothutu = 1;
    }

    $objPHPExcel = new PHPExcel();
    //set margin
    $sheet = $objPHPExcel->getActiveSheet();
    $pageMargins = $sheet->getPageMargins();
    $margin5 = 0.5 / 2.54;
    $margin8 = 0.8 / 2.54;
    $margin9 = 0.9 / 2.54;
    $margin012 = 0.12 / 2.54;
    $pageMargins->setTop($margin012);
    $pageMargins->setBottom($margin012);
    $pageMargins->setLeft($margin012);
    $pageMargins->setRight($margin012);
    // set border
    $styleArrayBorder = array(
        'borders' => array(
            'allborders' => array(
                'style' => PHPExcel_Style_Border::BORDER_DOTTED
            )
        )
    );
    $objPHPExcel->getDefaultStyle()->applyFromArray($styleArrayBorder);

    // set footer number
    $objPHPExcel->getActiveSheet()
                ->getHeaderFooter()
                ->setOddFooter('Page &P');
    $objPHPExcel->getActiveSheet()
                ->getHeaderFooter()
                ->setEvenFooter('Page &P');

    //set A4 landscape
    $objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setFitToWidth(1);
    $objPHPExcel->getActiveSheet()->getPageSetup()->setFitToHeight(0);
    //set width, height
    $widthCell = 23.5;
    $objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth($widthCell);
    $objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth($widthCell);
    $objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth($widthCell);
    $objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth($widthCell);
    $objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth($widthCell);
    $objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth($widthCell);
    //set align
    $objPHPExcel->getActiveSheet()
                ->getStyle('A')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()
                ->getStyle('B')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()
                ->getStyle('C')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()
                ->getStyle('D')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()
                ->getStyle('E')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);
    $objPHPExcel->getActiveSheet()
                ->getStyle('F')
                ->getAlignment()
                ->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER)
                ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

    //export
    $stt = 0;
    $sizeWrapLine = 11;
    $sizeText = 12;
    foreach($ketqua as $k=>$v){
        $objPHPExcel->getActiveSheet()->getRowDimension($k)->setRowHeight(95);
        if(isset($v[1])) {
            $objPHPExcel->getActiveSheet()->getStyle('A')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[1]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[1]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[1]['size']."#       x      ".$v[1]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('A'.$k, $objRichText); 
        }
        if(isset($v[2])) {
            $objPHPExcel->getActiveSheet()->getStyle('B')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[2]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[2]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[2]['size']."#       x      ".$v[2]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('B'.$k, $objRichText);
        }
        if(isset($v[3])) {
            $objPHPExcel->getActiveSheet()->getStyle('C')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[3]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[3]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[3]['size']."#       x      ".$v[3]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('C'.$k, $objRichText);
        }
        if(isset($v[4])) {
            $objPHPExcel->getActiveSheet()->getStyle('D')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[4]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[4]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[4]['size']."#       x      ".$v[4]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('D'.$k, $objRichText);
        }
        if(isset($v[5])) {
            $objPHPExcel->getActiveSheet()->getStyle('E')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[5]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[5]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[5]['size']."#       x      ".$v[5]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('E'.$k, $objRichText);
        }
        if(isset($v[6])) {
            $objPHPExcel->getActiveSheet()->getStyle('F')->getAlignment()->setWrapText(true);
            $objRichText = new PHPExcel_RichText();
            $objsize1 = $objRichText->createTextRun($v[6]['tenhang']."\n");
            $objsize1->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize2 = $objRichText->createTextRun($v[6]['mahang']."\n");
            $objsize2->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $warpline = $objRichText->createTextRun(" \n");
            $warpline->getFont()->setSize($sizeWrapLine);
            $objsize3 = $objRichText->createTextRun($v[6]['size']."#       x      ".$v[6]['soluong']);
            $objsize3->getFont()->setSize($sizeText)->setBold(true)->setName('VNI-Souvir');
            $objPHPExcel->getActiveSheet()->setCellValue('F'.$k, $objRichText);
        }
    }
    // Rename sheet 
    $objPHPExcel->getActiveSheet()->setTitle('Tem Hang'); 
    // Set active sheet index to the first sheet, so Excel opens this as the first sheet 
    $objPHPExcel->setActiveSheetIndex(0); 
    // Redirect output to a client’s web browser (Excel5) 
    header('Content-Type: application/vnd.ms-excel'); 
    header('Content-Disposition: attachment;filename="temhang.xls"'); 
    header('Cache-Control: max-age=0'); 
        
    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5'); 
    $objWriter->save('php://output'); 
    exit;
}

?>