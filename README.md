# ank-phpexcel
## 导出excel使用方法
首先查询出要导出的列表数据

```
//从数据库中查询出数据设置给data;
$conf=[
        'data'     => [
            ['id' => 123, 'title' => '标题'],
        ],
        //下载的文件名字不用加后缀
        'filename' => 'data',
        //定义excel中要使用的数据字段和字段对应的标题
        'field'    =>[
            'id' => '序号',
            'title' => '文章标题',
            '#link#url'=>'http://www.baidu.com',
            '#pic#pic'=>'./1.png'
        ],
    ];
```
>id,title这些字段前面可以添加一些特殊格式的标识 **#标识#** 来插入对应的特殊数据  
如 #pic#id 插入图片  id的值会被识别为图片的路径  
 #link#url插入链接  url的值为url链接
## 导入excel数据的方法
第一行导入后会默认做为字段，真实的数据是从第二行开始的
格式为:
```
title   name    price
标题1   名字1   130
标题2   名字2   140
```
导入后的格式为一个二维数组:
```
array(
array('title'=>"标题1",'name'=>'名字1','price'=>'130'),
array('title'=>"标题2",'name'=>'名字1','price'=>'140'),
)
```
首先建立一个表单
```
<form action="" method="post">
<input type="file" name="excel">
<input type="submit" value="导入">
</form>
```
//使用方法
```
$excel=new \ank\Excel();
//两个参数一个是上传后保存的路径另一个是上传的表单name,如果第一个路径文件存在就直接解析
$excel->importExcel('e:/upload','excel');
//参数是一个配置项如上面导出的配置一样
$excel->downloadExcel($conf);
//更多用法可以参考example.php中的使用方法
```