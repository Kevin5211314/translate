<?php

/**
 * 这是一个支持word文档转string
 * Class Docx2text
 * @package dingjingfei\Docx2text
 */
namespace Kevin5211314\translate;

class Docx2text
{
    
 	const SEPARATOR_TAB = "\t";

    const SEPARATOR_BR = "<br/>";

    /**
     * 对象  zipArchive
     *
     * @var string
     * @access private
     */
    private $docx;

    /**
     * document.xml中的对象domDocument
     *
     * @var string
     * @access private
     */
    private $domDocument;

    /**
     * xml from document.xml
     *
     * @var string
     * @access private
     */
    private $_document;

    /**
     * xml from document.xml
     *
     * @var string
     * @access private
     */
    private $_newDocument;


    /**
     * xml from numbering.xml
     *
     * @var string
     * @access private
     */
    private $_numbering;

    /**
     *  xml 从 footnote
     *
     * @var string
     * @access private
     */
    private $_footnote;

    /**
     *  xml 从 endnote
     *
     * @var string
     * @access private
     */
    private $_endnote;

    /**
     *文档所有尾注的数组
     *
     * @var string
     * @access private
     */
    private $endnotes;

    /**
     * 文档所有脚注的数组
     *
     * @var string
     * @access private
     */
    private $footnotes;

    /**
     * 数组的所有关系的文件
     *
     * @var string
     * @access private
     */
    private $relations;

    /**
     * 像列表一样插入字符数组
     *
     * @var string
     * @access private
     */
    private $numberingList;

    /**
     * 将被导出的文本内容
     *
     * @var string
     * @access private
     */
    private $textOuput;


    /**
     * 布尔变量知道图表是否将被转换为文本
     *
     * @var string
     * @access private
     */
    private $chart2text;

    /**
     * 布尔变量知道表是否将被转换为文本
     *
     * @var string
     * @access private
     */
    private $table2text;

    /**
     * 布尔变量知道列表是否将被转换为文本
     *
     * @var string
     * @access private
     */
    private $list2text;

    /**
     * 布尔变量知道段落是否会被转换为文本
     *
     * @var string
     * @access private
     */
    private $paragraph2text;

    /**
     * 布尔变量知道脚注是否会被取消
     *
     * @var string
     * @access private
     */
    private $footnote2text;

    /**
     * 布尔变量知道是否提取尾注
     *
     * @var string
     * @access private
     */
    private $endnote2text;

    /**
     * 布尔变量知道是否提取尾注
     *
     * @var string
     * @access private
     */
    private $media;

    private $fileName;

    /**
     * 所有的图片路径信息
     * Enter description here ...
     * @var unknown_type
     */
    private $allImgPaths;
    
    /**
     * Construct
     *
     * @param $boolTransforms布尔值数组应该被转换或不转换的数组
     * @access public
     */

    public function __construct($boolTransforms = array())
    {
        //table,list, paragraph, footnote, endnote, chart, media
        if (isset($boolTransforms['table'])) {
            $this->table2text = $boolTransforms['table'];
        } else {
            $this->table2text = true;
        }

        if (isset($boolTransforms['list'])) {
            $this->list2text = $boolTransforms['list'];
        } else {
            $this->list2text = true;
        }

        if (isset($boolTransforms['media'])) {
            $this->media2text = $boolTransforms['media'];
        } else {
            $this->media2text = true;
        }

        if (isset($boolTransforms['paragraph'])) {
            $this->paragraph2text = $boolTransforms['paragraph'];
        } else {
            $this->paragraph2text = true;
        }

        if (isset($boolTransforms['footnote'])) {
            $this->footnote2text = $boolTransforms['footnote'];
        } else {
            $this->footnote2text = true;
        }

        if (isset($boolTransforms['endnote'])) {
            $this->endnote2text = $boolTransforms['endnote'];
        } else {
            $this->endnote2text = true;
        }

        if (isset($boolTransforms['chart'])) {
            $this->chart2text = $boolTransforms['chart'];
        } else {
            $this->chart2text = true;
        }

        $this->textOuput = '';
        $this->media = '';
        $this->docx = '';
        $this->_numbering = '';
        $this->numberingList = array();
        $this->endnotes = array();
        $this->footnotes = array();
        $this->medianotes = array();
        $this->_media = array();
        $this->relations = array();
    }
    
    /**
     * 遍历图片节点信息
     * Enter description here ...
     * @param unknown_type $node
     * @param unknown_type $retArray
     * @param unknown_type $imgWHArray
     */
    public function vistNodes($node,&$retArray,&$imgWHArray){
        if($node == null){
            return;
        }
        if($node->nodeName=='wp:extent'){
        	if($node->hasAttributes()){
        		 $length = $node->attributes->length;
        		 $width = 0;
        		 $height = 0;
        		 for ($i = 0; $i < $length; ++$i){
        		 	$item = $node->attributes->item($i);
        		 	if($item->nodeName == "cx"){
        		 		$width = round($item->nodeValue*100/914400);
        		 	}else if($item->nodeName == "cy"){
        		 		$height = round($item->nodeValue*100/914400);
        		 	}
        		 }
        		 $info = array('width'=>$width,'height'=>$height);
        		 $imgWHArray[] = $info;
        	}
        }else if($node->nodeName=='v:shape'){
        	if($node->hasAttributes()){
        		$length = $node->attributes->length;
        		$width = 0;
        		$height = 0;
        		for ($i = 0; $i < $length; ++$i){
        			$item = $node->attributes->item($i);
        			if($item->nodeName == "style"){
        				$style = $item->nodeValue;
        				$data = explode(";",$style);
                        foreach ($data as $key => $value) {
                           $data[$key] = explode(":",$value); 
                        }
                        $infoarray = '';
                        foreach ($data as $key => $value) {
                            if ( $value[0] === '') {
                                unset($data[$key]);
                            }else{
                                 $infoarray[$key] = [ $value[0] => $value[1] ];
                            }
                        }
                        $height = '';
                        $width  = '';
                        foreach ($infoarray as $key => $value) {
                            $keys =  array_keys($value);
                            foreach ($keys as $k => $v) {
                                if ($v === 'height') {
                                    if ( strpos($value['height'], 'px')) {
                                        $height =  strstr($value['height'], 'px', TRUE);
                                    }else{
                                        $height =  strstr($value['height'], 'pt', TRUE);
                                        $height =  $height*4/3;
                                    }
                                }elseif ($v === 'width') {
                                    if ( strpos($value['width'], 'px')) {
                                        $width =  strstr($value['width'], 'px', TRUE);
                                    }else{
                                        $width =  strstr($value['width'], 'pt', TRUE);
                                        $width =  $width*4/3;
                                    }
                                }
                            }
                        }

        			}
        		}
        		$width = $width; //英寸转像素
        		$height = $height;
        	}
        	$info = array('width'=>$width,'height'=>$height);
        	$imgWHArray[] = $info;
        }
        if($node->nodeName=='a:blip'  or  $node->nodeName=='v:imagedata' ){
           if($node->hasAttributes()){
                $length = $node->attributes->length;
                for ($i = 0; $i < $length; ++$i) {
                	$item = $node->attributes->item($i);
                    if($item->nodeName == "r:embed"){
                		$retArray[] = $item->nodeValue;	
                	}else if($item->nodeName == "r:id"){
                		$retArray[] = $item->nodeValue;
                	}
                }   
                return $retArray;
            }
        }
        if($node->childNodes){
            foreach ($node->childNodes as $child){
                $this->vistNodes($child,$retArray,$imgWHArray);
            }
        }
    }

    /**
     *
     * 提取单词文档的内容，并在给定名称时创建文本文件。
     * @access public
     * @param string $filename of the word document.
     * @return string
     */
    public function extract()
    {   
        if (empty($this->_document)) {
            exit('There is no content');
        }
        $this->allImgPaths = $this->getAllImageNodes();
        $this->domDocument = new \DomDocument();
        $this->domDocument->loadXML($this->_document);
        //获取身体节点检查所有子节点的内容
        $bodyNode = $this->domDocument->getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'body');
        $bodyNode = $bodyNode->item(0);
        foreach ($bodyNode->childNodes as $child) {
            if ($this->table2text && $child->tagName == 'w:tbl') {
                $this->textOuput .= $this->table($child) . self::SEPARATOR_BR;
            } else {
                $this->textOuput .= $this->printWP($child) . self::SEPARATOR_BR;
            }
        }
        unset($bodyNode);
        return $this->textOuput;
    }

    /**
     * 提取w:p标签的内容
     * @access private
     * @param  节点对象 object
     * @return string
     */
    private function printWP($node)
    {   
        if ($this->list2text) {//将ooxml中的列表转换为带有制表符和项目符号的格式化列表
            if (empty($this->numberingList)) {//检查是否从zip压缩文件中提取numbering.xml
                $this->listNumbering();
            }
        }
        $text = '';
        foreach ($node->childNodes as $child){
        	if ( $child->tagName != 'w:instrText') {
        		if ($child->hasChildNodes()) {
        			foreach ($child->childNodes as $node) {
        				if ( $node->tagName === 'w:instrText') {
        					$node->nodeValue = '';
        				}
        			}
        		}
                $text .= $this->toText($child);  //获取全部的text  
        	}
        }
        return $text;
    }

    /**
     * 处理所有的图片节点，查看用户上传的word中的全部的图片
     * Enter description here ...
     * @param $id
     */
    public function getAllImageNodes(){
    	$imgDocx = new \ZipArchive();
    	$ret = $imgDocx->open($this->fileName);
    	if ($ret === true) {
    		$media = $imgDocx->getFromName('word/_rels/document.xml.rels');
    		$domDocument = new \DomDocument();
    		$domDocument->loadXML($media);
    		$retArray = array();
    		foreach ($domDocument->childNodes as $child) {
    			if($child->nodeName  == 'Relationships'){
    				foreach ($child->childNodes as $node){
    					if($node->nodeName=='Relationship'){
    						if($node->hasAttributes()){
    							$length = $node->attributes->length;
    							for ($i = 0; $i < $length; ++$i) {
    								$item = $node->attributes->item($i);
    								$retArray[] = array($item->nodeName=>$item->nodeValue);
    							}
    						}
    					}
    				}
    			}
    		}
    		unset($domDocument);
    		$imgDocx->close();
    		return $retArray;
		}else{
			$imgDocx->close();
		}
	}

	/**
	 * 如果题目下有图片，则处理图片
	 * @param    [int]      $id        [图片的id]
	 * @param    [string]   $filepath  [文件路径]
	 * @return   [string]   $path      图片路径
	 * @Author   Kevin.D
	 * @DateTime 2017-12-26
	 */
	public function ProcessingPictures($id){
		foreach ($this->allImgPaths as $key => $value) {
			if ( $value['Id'] == $id ) {
				return $this->PictureFinal($this->allImgPaths[$key+2]['Target']);
			}
		}
		return null;
	}
    
	/**
	 * 查找具体的图片  getPictureFinal
     *
	 * @author   Kevin.D
	 * @access   public
	 * @param    $imagename    [图片名称]
	 */
	public function getPictureFinal( $imagename ){
		$medias = $this->docx->getFromName('word/'.$imagename);
		$base_img = base64_encode($medias);
		$imagename = substr($imagename,7);
		$reply['imgname'] = $imagename;
		$reply['base64'] = $base_img;
		return $reply;
	}

    /**
     * Setter
     *
     * @access public
     * @param $filename
     */
    public function setDocx($filename)
    {   
    	$this->fileName = $filename;
        $this->docx = new \ZipArchive();
        $ret = $this->docx->open($filename);
        if ($ret === true) {
            $this->_document = $this->docx->getFromName('word/document.xml');
            return true;
        } else {
        	$this->docx->close();
            return false;
        }
    }
    
    /**
     * 关闭
     * Enter description here ...
     */
    public function closeDocx(){
    	if($this->docx != null){
    		$this->docx->close();
    	}
    	unset($this->domDocument);
    }

    /**
     * 将内容从endnote.xml中提取到数组中
     *
     * @access private
     */
    private function loadEndNote()
    {
        if (empty($this->endnotes)) {
            if (empty($this->_endnote)) {
                $this->_endnote = $this->docx->getFromName('word/endnotes.xml');
            }
            if (!empty($this->_endnote)) {
                $domDocument = new DomDocument();
                $domDocument->loadXML($this->_endnote);
                $endnotes = $domDocument->getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'endnote');
                foreach ($endnotes as $endnote) {
                    $xml = $endnote->ownerDocument->saveXML($endnote);
                    $this->endnotes[$endnote->getAttribute('w:id')] = trim(strip_tags($xml));
                }
                unset($domDocument);
            }
        }
    }

    /**
     * 从footnote.xml中将内容提取到数组
     *
     * @access private
     */
    private function loadFootNote()
    {
        if (empty($this->footnotes)) {
            if (empty($this->_footnote)) {
                $this->_footnote = $this->docx->getFromName('word/footnotes.xml');
            }
            if (!empty($this->_footnote)) {
                $domDocument = new DomDocument();
                $domDocument->loadXML($this->_footnote);
                $footnotes = $domDocument->getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'footnote');
                foreach ($footnotes as $footnote) {
                    $xml = $footnote->ownerDocument->saveXML($footnote);
                    $this->footnotes[$footnote->getAttribute('w:id')] = trim(strip_tags($xml));
                }
                unset($domDocument);
            }
        }
    }

    /**
     * 将列表中的图片提取到一个数组
     *
     * @access private
     */
    private function loadMediaNote()
    {
        $ids = array();
        $nums = array();
        //从zip归档中获取xml代码
        $this->_numbering = $this->docx->getFromName('word/numbering.xml');
        if (!empty($this->_numbering)) {
            //我们使用domdocument迭代编号标签的子代
            $domDocument = new \DomDocument();
            $domDocument->loadXML($this->_numbering);
            $numberings = $domDocument->getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'numbering');
            //numbering.xml中只有一个编号标记
            $numberings = $numberings->item(0);
            foreach ($numberings->childNodes as $child) {
                $flag = true;//布尔变量知道节点是否是列表的第一个样式
                foreach ($child->childNodes as $son) {
                    if ($child->tagName == 'w:abstractNum' && $son->tagName == 'w:lvl') {
                        foreach ($son->childNodes as $daughter) {
                            if ($daughter->tagName == 'w:numFmt' && $flag) {
                                $nums[$child->getAttribute('w:abstractNumId')] = $daughter->getAttribute('w:val');//设置列表的内部索引的键和值是它的类型的项目符号
                                $flag = false;
                            }
                        }
                    } elseif ($child->tagName == 'w:num' && $son->tagName == 'w:abstractNumId') {
                        $ids[$son->getAttribute('w:val')] = $child->getAttribute('w:numId');//$ids是列表的索引
                    }
                }
            }
            //一旦我们知道文件中有什么样的列表，就准备好了图书馆将要使用的子弹
            foreach ($ids as $ind => $id) {
                if ($nums[$ind] == 'decimal') {
                    //如果类型是十进制，则意味着子弹将是数字
                    $this->numberingList[$id][0] = range(1, 10);
                    $this->numberingList[$id][1] = range(1, 10);
                    $this->numberingList[$id][2] = range(1, 10);
                    $this->numberingList[$id][3] = range(1, 10);
                } else {
                    //否则是*和其他字符
                    $this->numberingList[$id][0] = array('*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*');
                    $this->numberingList[$id][1] = array(chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175));
                    $this->numberingList[$id][2] = array(chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237));
                    $this->numberingList[$id][3] = array(chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248));
                }
            }
            unset($domDocument);
        }
    }

    /**
     * 将列表的样式提取到一个数组
     *
     * @access private
     */
    private function listNumbering()
    {
        $ids = array();
        $nums = array();
        //从zip归档中获取xml代码
        $this->_numbering = $this->docx->getFromName('word/numbering.xml');
        if (!empty($this->_numbering)) {
            //我们使用domdocument迭代编号标签的子代
            $domDocument = new \DomDocument();
            $domDocument->loadXML($this->_numbering);
            $numberings = $domDocument->getElementsByTagNameNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'numbering');
            //numbering.xml中只有一个编号标记
            $numberings = $numberings->item(0);
            foreach ($numberings->childNodes as $child) {
                $flag = true;//布尔变量知道节点是否是列表的第一个样式
                foreach ($child->childNodes as $son) {
                    if ($child->tagName == 'w:abstractNum' && $son->tagName == 'w:lvl') {
                        foreach ($son->childNodes as $daughter) {
                            if ($daughter->tagName == 'w:numFmt' && $flag) {
                                $nums[$child->getAttribute('w:abstractNumId')] = $daughter->getAttribute('w:val');//设置列表的内部索引的键和值是它的类型的项目符号
                                $flag = false;
                            }
                        }
                    } elseif ($child->tagName == 'w:num' && $son->tagName == 'w:abstractNumId') {
                        $ids[$son->getAttribute('w:val')] = $child->getAttribute('w:numId');//$ids是列表的索引
                    }
                }
            }
            //一旦我们知道文件中有什么样的列表，就准备好了图书馆将要使用的子弹
            foreach ($ids as $ind => $id) {
                if ($nums[$ind] == 'decimal') {
                    //如果类型是十进制，则意味着子弹将是数字
                    $this->numberingList[$id][0] = range(1, 10);
                    $this->numberingList[$id][1] = range(1, 10);
                    $this->numberingList[$id][2] = range(1, 10);
                    $this->numberingList[$id][3] = range(1, 10);
                } else {
                    //否则是*和其他字符
                    $this->numberingList[$id][0] = array('*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*', '*');
                    $this->numberingList[$id][1] = array(chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175), chr(175));
                    $this->numberingList[$id][2] = array(chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237), chr(237));
                    $this->numberingList[$id][3] = array(chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248), chr(248));
                }
            }
            unset($domDocument);
        }
    }

    /**
     *
     * 从document.xml中提取表节点的内容并返回文本内容
     *
     * @access private
     * @param $node object
     *
     * @return string
     */
    private function table($node)
    {
        $output = '';
        if ($node->hasChildNodes()) {
            foreach ($node->childNodes as $child) {
                //开始表格的新行
                if ($child->tagName == 'w:tr') {
                    foreach ($child->childNodes as $cell) {
                        //开始一个新的细胞
                        if ($cell->tagName == 'w:tc') {
                            if ($cell->hasChildNodes()) {
                                foreach ($cell->childNodes as $p) {
                                    $output .= $this->printWP($p);
                                }
                                $output .= self::SEPARATOR_TAB;
                            }
                        }
                    }
                }
                $output .= self::SEPARATOR_BR;
            }
        }
        return $output;
    }

    /**
     * 从document.xml中提取节点的内容，并只返回文本内容和。 剥离html标签
     * @access private
     * @param $node object
     * @return string
     */
    private function toText($node)
    {
        if($node != null ){
            if($node->childNodes != null){
                $fonts = "";
                $lang = "";
                $outputText = "";
                foreach ($node->childNodes as $child){
                    if($child->tagName == "w:rPr"){
                        foreach($child->childNodes as $suChild){
                            if($suChild->tagName == 'w:rFonts'){
                                $fonts = $suChild->getAttribute('w:eastAsia');
                                if($fonts == ""){
                                    $fonts = $suChild->getAttribute('w:eastAsiaTheme');
                                }
                                if($fonts == ""){
                                    $fonts = $suChild->getAttribute('w:ascii');
                                }
                            }else if($suChild->tagName == 'w:lang'){
                                $lang = $suChild->getAttribute('w:eastAsia');
                            }
                        }
                    }else{
                        $outputText .= $child->ownerDocument->saveXML($child);
                    }
                }
                $outputText = strip_tags($outputText);
                if($outputText == false || $outputText === "false" || $outputText == ""){
                    return "";
                }
                $fonts = '宋体'; // 无论有没有字体样式，全部使用宋体，全部的字体全部使用宋体
                if($lang != ""){
                    return "<span lang='{$lang}' style='font-family:{$fonts}'>{$outputText}</span>";
                }else{
                    return "<span style='font-family:{$fonts}'>{$outputText}</span>";
                }
                return "<span style='font-family:{$fonts}'>{$outputText}</span>";
            }else{
                $xmllist = $node->ownerDocument->saveXML($node);
                // if($xmllist == false || $xmllist === "false"){
                //     return "";
                // }
                // $xml = ;
                // if($xml === "false"){
                //     return "";
                // }
                return strip_tags($xmllist);
            }
        }else{
            return "";
        }
    }

    /**
     * 删除空格 删除全部的特殊字符
     * @Author    Kevin.D
     * @DateTime  2017-12-28
     * @return    string       $str
     */
    public function trimall($str)
    {
        $oldchar=array("　","\t","\n","\r");
        $newchar=array("","","<br>","");
        return str_replace($oldchar,$newchar,$str);
    }
    
}