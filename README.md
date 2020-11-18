#  Docx2text
## 这是一个支持word文档转string

- v1.0
	- 命中一个后返回
- v2.0
	- 支持命中多个返回
	- 支持在树梢增加自定义数组 [替换内容] 
	- 性能提升10倍
- v3.0
    - 增加删除特性
        - 删除整棵关键词树
    - 解决命中不全BUG
    - 3.1
        - 增加词频统计
    - 3.5
        - 清除词频统计 [没有什么意义]
        - 增加Suggestion特性  根据某个word提取相关的词语
            - 所有检索依据字典
            - 提取关联词均为从左至右原则
            - 因为个人更倾向其为一个“组件服务”，所以增加拼音索引需要主动增加

## 使用场景
- 文档识别
- 文档处理
- 数据导入

## 注意
- 需要严格注意文档格式，暂不支持excel。

## 示例
```php
<?php
require "../src/TranslateWord.php";

## new对象
$docx2txt  = new Kevin5211314\translate\Docx2text();

## 设置docx文档
$result    = $docx2txt->setDocx('./Examples.docx');

## 开始处理导出
$replyText = $docx2txt->extract();

## 输出结果
echo $text;

```

## composer安装
```
composer require kevinfei/translats
```

