<?php
include __dir__ . '/libs/Excel.php';
$excel = new \ank\Excel();
//两个参数一个是上传后保存的路径另一个是上传的表单name,如果第一个路径文件存在就直接解析
$arr = $excel->importExcel(__dir__ . '/test.xls', 'excel');
var_dump($arr);
die();
//参数是一个配置项如上面导出的配置一样
$conf = [
    'data'     => [
        ['id' => 123, 'title' => '标题'],
        ['id' => 123, 'title' => '标题'],
        ['id' => 123, 'title' => '标题'],
        ['id' => 123, 'title' => '标题'],
        ['id' => 123, 'title' => '标题'],

    ],
    //下载的文件名字不用加后缀
    'filename' => 'data',
    //定义excel中要使用的数据字段和字段对应的标题
    'field'    => [
        'id'    => '序号',
        'title' => '文章标题',
        // '#link#url' => 'http://www.baidu.com',
        // '#pic#pic'  => './1.png',
    ],
];
$excel->downloadExcel($conf, __dir__ . '/1.xls');
