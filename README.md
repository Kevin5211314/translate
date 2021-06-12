#  Docx2text
## 这是一个支持word文档转string

- v1.0
	- 读取后返回string
- v2.0
	
- v3.0
    

## 使用场景
- 文档识别
- 文档处理
- 数据导入

## 注意
- 需要严格注意文档格式，暂不支持excel。

## 示例
```php
<?php
require "../src/Docx2text.php";

## new对象
$docx2txt  = new kevinfei\translate\Docx2text();

## 设置docx文档
$result    = $docx2txt->setDocx('./Examples.docx');

## 开始处理导出
$replyText = $docx2txt->extract();

## 输出结果
echo $replyText;
```

## composer安装
```
composer require kevinfei/translate
```

