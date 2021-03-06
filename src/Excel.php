<?php

namespace ank;

use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 *使用方法
 *$field="'diqu,goods_id,num,create_time";
 *$fieldtitle="市场,产品名字,数量,下单时间";
 *$str=run_plugin_method('Phpexcel','export',array($list,
 *'shichang-tongji',
 *array($field,$fieldtitle)
 *)
 *);
 ****/
// require_once path_a('/Plugins/Plugin.class.php');
class Excel
{

    /**
     * 导入excel文件中的数据
     * @param null   $excelFilePath
     * @param string $formname
     * @return array
     */
    public function importExcel($excelFilePath = null, $formname = 'excel')
    {
        try {
            $uploadfile = $excelFilePath;
            if (!$uploadfile) {
                if ($_FILES && isset($_FILES[$formname]['name'])) {
                    $file       = $_FILES[$formname]['name'];
                    $uploadfile = $_FILES[$formname]['tmp_name'];
                }
            }
            //$isuploaded = true;
            //if (is_file($uploadpath)) {
            //    $isuploaded = false;
            //}
            //$result     = false;
            //$uploadfile = '';
            //if ($isuploaded) {
            //$file         = '';
            //$filetempname = '';

            if (!is_file($uploadfile)) {
                return [];
            }
            //自己设置的上传文件存放路径
            //$filePath = $uploadpath;
            //$str      = "";

            //注意设置时区
            //$time = date("y-m-d-H-i-s"); //去当前上传的时间
            //获取上传文件的扩展名
            //$extend = strrchr($file, '.');
            //上传后的文件名
            //$name       = $time . $extend;
            //$uploadfile = $filePath . '/' . $name;                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            //上传后的文件名地址
            //move_uploaded_file() 函数将上传的文件移动到新位置。若成功，则返回 true，否则返回 false。
            //$result     = move_uploaded_file($filetempname, $uploadfile);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         //假如上传到当前目录下
            //$uploadfile = $filetempname;
            //$result     = true;
            //}
            //else {
            //    $result     = true;
            //    $uploadfile = $uploadpath;
            //}
            //echo $result;
            //if ($result) //如果上传文件成功，就执行导入excel操作
            //{
            // include "conn.php";
            // $objReader   = \PHPExcel_IOFactory::createReader('Excel5'); //use excel2007 for 2007 format
            // $objPHPExcel = $objReader->load($uploadfile);

            $extension   = strtolower(pathinfo($uploadfile, PATHINFO_EXTENSION));
            $objPHPExcel = \PhpOffice\PhpSpreadsheet\IOFactory::load($uploadfile);
            //if ($extension == 'xlsx') {
            //    //$objReader   = new \PHPExcel_Reader_Excel2007();
            //    //$objPHPExcel = $objReader->load($uploadfile);
            //    $objPHPExcel = \PhpOffice\PhpSpreadsheet\IOFactory::load("05featuredemo.xlsx");
            //}
            //elseif ($extension == 'xls') {
            //    $objReader   = new \PHPExcel_Reader_Excel5();
            //    $objPHPExcel = $objReader->load($uploadfile);
            //}
            //elseif ($extension == 'csv') {
            //    $PHPReader = new \PHPExcel_Reader_CSV();
            //    //默认输入字符集
            //    $PHPReader->setInputEncoding('GBK');
            //    //默认的分隔符
            //    $PHPReader->setDelimiter(',');
            //    //载入文件
            //    $objPHPExcel = $PHPReader->load($uploadfile);
            //}

            $sheet         = $objPHPExcel->getSheet(0);
            $highestRow    = $sheet->getHighestRow();       //取得总行数
            $highestColumn = $sheet->getHighestColumn();    //取得总列数
            //dump($highestRow);
            //dump($highestColumn);
            $colarr   = [];
            $fieldarr = [];
            //循环读取excel文件,读取一条,插入一条
            for ($j = 1; $j <= $highestRow; $j++) //从第一行开始读取数据
            {

                $tem = [];
                for ($k = 'A'; $k <= $highestColumn; $k++) //从A列读取数据
                {

                    //这种方法简单，但有不妥，以'\\'合并为数组，再分割\\为字段值插入到数据库
                    //实测在excel中，如果某单元格的值包含了\\导入的数据会为空
                    //getValue 取内容   getCalculatedValue 取工式结果
                    $str = $objPHPExcel->getActiveSheet()->getCell("$k$j")->getCalculatedValue(); //读取单元格
                    $str = trim($str);
                    //第一行做为字段
                    if (empty($str) && $j == 1) {
                        break;
                    }

                    if ($j == 1) {
                        $fieldarr[$k] = $str;
                    }
                    else {
                        //遍历的列数据超过啦字段长度直接退出
                        if (!isset($fieldarr[$k])) {
                            break;
                        }
                        //如果第一列数据为空,也视为结束
                        if ($k == 'A' && empty($str)) {
                            break;
                        }
                        $tem[$fieldarr[$k]] = $str;
                    }
                }
                //非第一行做为数据填充
                if ($j !== 1 && $tem) {
                    $colarr[] = $tem;
                }

            }
            //$isuploaded && unlink($uploadfile); //删除上传的excel文件
            return $colarr;
            //}
            //else {
            //    return [];
            //}
        } catch (\Exception $e) {
            return [];
        }
    }

    public function downloadExcel($config = [], $savepath = null)
    {
        $conf = [
            'data'     => [], //数据
            'filename' => '', //文件名字
            'field'    => '', //excel标题名字
        ];
        // dump($config);
        $conf = array_merge($conf, $config);
        $data = $conf['data'];

        $fileName   = $conf['filename'] . '-' . date('Y-m-d H-i-s');
        $field      = array_keys($conf['field']);
        $fieldtitle = array_values($conf['field']);

        $spreadsheet = new Spreadsheet();
        $sheet       = $spreadsheet->getActiveSheet();
        //表头
        //设置单元格内容
        $titCol = 'A';
        foreach ($conf['field'] as $key => $value) {
            // 单元格内容写入
            $sheet->setCellValue($titCol . '1', $value);
            $titCol++;
        }
        $row = 2;                                                      // 从第二行开始
        foreach ($data as $item) {
            $dataCol = 'A';
            foreach ($conf['field'] as $key => $value) {
                // 单元格内容写入
                $value = ($item[$key] ?? '');
                if (is_numeric($value)) {
                    $sheet->getStyle($dataCol . $row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
                }
                $sheet->setCellValue($dataCol . $row, $value);
                $dataCol++;
            }
            $row++;
        }
        if ($savepath !== null) {
            // Save
            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');

            $writer->save($savepath . '/' . $fileName . '.xlsx');

        }
        else {
            // Redirect output to a client’s web browser (Xlsx)
            header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            header('Content-Disposition: attachment;filename="' . $fileName . '.xlsx"');
            header('Cache-Control: max-age=0');
            // If you're serving to IE 9, then the following may be needed
            header('Cache-Control: max-age=1');

            // If you're serving to IE over SSL, then the following may be needed
            header('Expires: Mon, 26 Jul 1997 05:00:00 GMT');              // Date in the past
            header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
            header('Cache-Control: cache, must-revalidate');               // HTTP/1.1
            header('Pragma: public');                                      // HTTP/1.0

            $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
            $writer->save('php://output');
        }
        exit;
    }

    /**
     * 导出excel文件
     * @param array $config [description]
     * @return [type]         [description]
     */
    public function downloadExcel2($config = [], $savepath = '')
    {
        $conf = [
            'data'     => [], //数据
            'filename' => '', //文件名字
            'field'    => '', //excel标题名字
        ];
        // dump($config);
        $conf = array_merge($conf, $config);
        $data = $conf['data'];
        $name = $conf['filename'] . '-' . date('Y-m-d H-i-s');
        // dump($conf);
        // die();
        // $field      = explode(',', $conf['field'][0]);
        // $fieldtitle = explode(',', $conf['field'][1]);
        $field       = array_keys($conf['field']);
        $fieldtitle  = array_values($conf['field']);
        $objPHPExcel = new \PHPExcel();
        //以下是一些设置 ，什么作者  标题啊之类的
        $objPHPExcel->getProperties()->setCreator("ainiku")
                    ->setLastModifiedBy("ainiku")
                    ->setTitle("feilv export")
                    ->setSubject("feilv export")
                    ->setDescription("bakdata")
                    ->setKeywords("excel")
                    ->setCategory("result file");
        //设置表头
        $obj = $objPHPExcel->setActiveSheetIndex(0);

        $fieldnum = count($fieldtitle);
        $j        = 65;
        $pre_zm   = '';
        foreach ($fieldtitle as $v) {
            $zm = chr($j++);
            $obj->setCellValue($pre_zm . $zm . '1', ' ' . $v);
            if ($zm == 'Z') {
                $pre_zm .= 'A';
                $j      = 65;
            }
        }
        //Set border colors 设置边框颜色
        //$obj->freezePane(chr(65).'1:'.chr($j).'1');
        $objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getLeft()->getColor()->setARGB('FF993300');
        $objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getTop()->getColor()->setARGB('FF993300');
        $objPHPExcel->getActiveSheet()->getStyle('D13')->getBorders()->getBottom()->getColor()->setARGB('FF993300');
        $objPHPExcel->getActiveSheet()->getStyle('E13')->getBorders()->getRight()->getColor()->setARGB('FF993300');
        //Set border colors 设置背景颜色
        //$objPHPExcel->getActiveSheet()->getStyle(chr(65).'1:'.chr($j).'1')->getFill()->setFillType(\PHPExcel_Style_Fill::FILL_SOLID);
        //$objPHPExcel->getActiveSheet()->getStyle(chr(65).'1:'.chr($j).'1')->getFill()->getStartColor()->setARGB('FFededed');
        //以下就是对处理Excel里的数据， 横着取数据，主要是这一步，其他基本都不要改
        // 固定第一行
        $obj->freezePane('A1');
        $fieldnum = count($field);
        foreach ($data as $k => $v) {
            $num      = $k + 2;
            $temfield = $field;
            $j        = 'A';
            $i        = 0;
            //$obj=$objPHPExcel->setActiveSheetIndex(0);
            for ($i; $i < $fieldnum; $i++) {
                $temstr = array_shift($temfield);
                if (substr($temstr, 0, 1) == "'") {
                    $temstr = str_replace("'", '', $temstr);
                    $obj->setCellValue($j . $num, ' ' . $v[$temstr]);
                }
                elseif (substr($temstr, 0, 5) == "#pic#") {
                    //插入图片
                    $temstr = str_replace("#pic#", '', $temstr);
                    $img    = new \PHPExcel_Worksheet_Drawing();
                    $img->setPath($v[$temstr]);          //写入图片路径
                    $img->setHeight(100);                //写入图片高度
                    $img->setWidth(100);                 //写入图片宽度
                    $img->setOffsetX(1);                 //写入图片在指定格中的X坐标值
                    $img->setOffsetY(1);                 //写入图片在指定格中的Y坐标值
                    $img->setRotation(1);                //设置旋转角度
                    $img->getShadow()->setVisible(true); //
                    $img->getShadow()->setDirection(50); //
                    $img->setCoordinates($j . $num);     //设置图片所在表格位置
                    //$objPHPExcel->getColumnDimension("$letter[$i]")->setWidth(20);
                    $obj->getDefaultRowDimension()->setRowHeight(100);
                    $img->setWorksheet($obj); //把图片写到当前的表格中

                    //$objActSheet->getCell('E26')->getHyperlink()->setUrl( 'http://www.phpexcel.net');    //超链接url地址
                    //$objActSheet->getCell('E26')->getHyperlink()->setTooltip( 'Navigate to website');  //鼠标移上去连接提示信息
                    //$obj->setCellValue($j.$num, $img);
                }
                elseif (substr($temstr, 0, 6) == "#link#") {
                    $temstr = str_replace("#link#", '', $temstr);
                    $obj->setCellValue($j . $num, $v[$temstr]);
                    $obj->getCell($j . $num)->getHyperlink()->setUrl($v[$temstr]);     //超链接url地址
                    $obj->getCell($j . $num)->getHyperlink()->setTooltip($v[$temstr]); //鼠标移上去连接提
                }
                else {
                    $obj->setCellValue($j . $num, ' ' . $v[$temstr]);
                }
                $j++;
            }
        }
        $objPHPExcel->getActiveSheet()->setTitle('User');
        $objPHPExcel->setActiveSheetIndex(0);
        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        if ($savepath) {
            $objWriter->save($savepath);
        }
        else {
            header('Content-Type: application/vnd.ms-excel');
            header('Content-Disposition: attachment;filename="' . $name . '.xls"');
            header('Cache-Control: max-age=0');
            $objWriter->save('php://output');
            exit;
        }

    }
}
