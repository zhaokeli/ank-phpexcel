<?php
include __dir__ . '/libs/Excel.php';
$excel = new \ank\Excel();
//参数一是excel文件路径如果存在则直接解析,参数二是上传的表单name参数一为null的情况下自动从文件域为excel的键读取文件内容
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
