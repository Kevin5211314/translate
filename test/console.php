<?php
/**
 *
 * Author: KevinFei
 * Date: 20/11/18
 * Time: 10:21
 */

##Examples

require "../src/TranslateWord.php";

## new对象
$docx2txt  = new kevinfei\translate\Docx2text();

## 设置docx文档
$result    = $docx2txt->setDocx('./Examples.docx');

## 开始处理导出
$replyText = $docx2txt->extract();

## 输出结果
echo $replyText;
