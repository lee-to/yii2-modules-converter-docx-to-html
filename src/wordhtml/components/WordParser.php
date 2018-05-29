<?php

namespace wordhtml\components;

use Yii;

/*CLASS LTDOCX2HTML convert .docx file to (x)html.
 *
 * created date: 20.03.2015
 * author: lee-to
 * e-mail: leetoplayaz@yandex.ru
 * version: 1.2
 * library: DOMDocument, ZipArchive
 * use examples:
 * $ltdocx = new LTDOCX2HTML();
 * $ltdocx->file = 'path_to_file';
 * $ltdocx->tmpDir = 'path_to_tmpdir';
 * $ltdocx->init();
 * $ltdocx->sendDownloadFile();
 * exit();
 *
 */
class WordParser {

    public $file;
    public $tmpDir = '/tmp';

    private $doc;
    private $zip;

    protected $output;
    protected $docVersion;

    function __construct() {
        $this->doc = new \DOMDocument();
        $this->zip = new \ZipArchive();
    }

    public function init(){
        $fileinfo = pathinfo($this->file);
        if($fileinfo['extension'] == 'docx'){
            $this->zip->open($this->file);

            $this->zip->extractTo($this->tmpDir);
            $this->zip->close();

            $this->doc->load($this->tmpDir.'/word/document.xml');

            $this->doc = $this->doc->documentElement;

        }
        else throw new \Exception('Необходим формат docx');

    }

    public function getVersion(){
        $this->docVersion = $this->doc->namespaceURI;

        return $this->docVersion;
    }

    private function extractRelXML(){
        $xmlFile = $this->tmpDir."/word/_rels/document.xml.rels";
        $xml = file_get_contents($xmlFile);
        if($xml == false){
            return false;
        }
        $xml = mb_convert_encoding($xml, 'UTF-8', mb_detect_encoding($xml));
        $parser = xml_parser_create('UTF-8');
        $data = array();
        xml_parse_into_struct($parser, $xml, $data);
        foreach($data as $value){
            if($value['tag']=="RELATIONSHIP"){
                if(isset($value['attributes']['TARGETMODE'])){
                    $rels[$value['attributes']['ID']] = array(0 => $value['attributes']['TARGET'], 3=> $value['attributes']['TARGETMODE']);
                } else {
                    $rels[$value['attributes']['ID']] = array(0 => $value['attributes']['TARGET']);
                }
            }
        }
        return $rels;
    }

    private function extractNumberingXML($liId){
        $xmlFile = $this->tmpDir."/word/numbering.xml";

        $doc = new \DOMDocument;
        $doc->load($xmlFile);

        $doc = $doc->documentElement;

        foreach($doc->getElementsByTagName('num') as $element){
            $numId = $element->attributes->getNamedItem('numId')->value;
            if($numId == $liId) {
                $absctractId = $element->getElementsByTagName('abstractNumId')->item(0)->attributes->item(0)->value;

                foreach($doc->getElementsByTagName('abstractNum') as $element){
                    //type of list
                    if($element->attributes->getNamedItem('abstractNumId')->value == $absctractId){
                        $numbType = $element->getElementsByTagName('numFmt')->item(0)->attributes->item(0)->value;
                        $type = ($numbType == 'bullet') ? 'ul' : 'ol';
                    }
                }
            }
        }

        return $type;
    }

    public function getHtml($htmlspecialchars=false, $generate_urls=false){
        $this->output = "";
        $listmark = false;
        $listdecimal = false;

        if ($this->doc->childNodes) {
            foreach ($this->doc->childNodes as $node) {

                if ($node->childNodes) {
                    foreach ($node->childNodes as $element) {
                        if($element->tagName=='w:tbl'){

                            $this->output .= "<table width=100% border=1>";
                            foreach ($element->childNodes as $elementTable) {
                                if ($elementTable->nodeName == 'w:tr') {
                                    $this->output .= "<tr>";
                                    foreach($elementTable->childNodes as $elementTableTd) {
                                        if($elementTable->nodeName=='w:tr') {
                                            $this->output.='<td>';

                                            $tag = array();
                                            $this->generateTags($elementTableTd, $tag, true);

                                            $this->output.= $tag['open'].$elementTableTd->nodeValue.$tag['close'];

                                            $this->output.='</td>';
                                        }
                                    }
                                    $this->output .= "</tr>";
                                }
                            }
                            $this->output .= "<table>";
                        }
                        elseif($element->tagName=='w:p') {

                            if ($this->hasElement($element, 'br') && $element->getElementsByTagName('br')->item(0)->getAttribute('w:type') == 'page') {
                                $this->output .= '<br data-type="page">';
                            }

                            //current element
                            $innerElements = ''; $tag = array();
                            //get current element attributes
                            $this->getElements($innerElements, $element, $element->nodeValue, $tag, array(), $generate_urls);

                            if($innerElements != '')
                            {
                                if(!isset($tag['open'])) {
                                    $tag['open'] = '';
                                }
                                if(!isset($tag['close'])) {
                                    $tag['close'] = '';
                                }

                                if($tag['open'] == '<li data-type="ul">'){
                                    if(!$listmark) {
                                        $this->output.='<ul>';
                                        $listmark=true;
                                    }
                                }
                                if($listmark){
                                    if($tag['open'] != '<li data-type="ul">'){
                                        $this->output .= '</ul>';
                                        $listmark = false;
                                    }
                                }

                                if($tag['open'] == '<li data-type="ol">'){
                                    if(!$listdecimal) {
                                        $this->output.='<ol>';
                                        $listdecimal=true;
                                    }
                                }
                                if($listdecimal){
                                    if($tag['open'] != '<li data-type="ol">'){
                                        $this->output .= '</ol>';
                                        $listdecimal = false;
                                    }
                                }

                                if(isset($tag['open']) && $tag['close']) $this->output .= $tag['open'].$innerElements.$tag['close'];
                                else $this->output .= '<p>'.$innerElements.'</p>';
                            }
                        }

                    }
                }
            }
        }

        $this->output = $this->reformatOutput($this->output);

        if($htmlspecialchars){
            return htmlspecialchars($this->output);
        }
        else {
            return $this->output;
        }

    }

    public function generateXls(){
        $this->getHtml(false);

        $trs = explode('<br data-type="page">', $this->output);

        $xls = "<html xmlns:x='urn:schemas-microsoft-com:office:excel'>
            <!--[if gte mso 9]>
            <xml>
                <x:ExcelWorkbook>
                    <x:ExcelWorksheets>
                        <x:ExcelWorksheet>
                            <x:Name>Sheet 1</x:Name>
                            <x:WorksheetOptions>
                                <x:Print>
                                    <x:ValidPrinterInfo/>
                                </x:Print>
                            </x:WorksheetOptions>
                        </x:ExcelWorksheet>
                    </x:ExcelWorksheets>
                </x:ExcelWorkbook>
            </xml>
            <![endif]-->
            <head>
                <meta charset='UTF-8' />
                <title></title>
            </head>
            <body>
            <table border='1' align='left'>
            <tr>
                <td bgcolor='#ffe4c4' width='100'>Ссылка</td>
                <td bgcolor='#ffe4c4' width='700'>Контент</td>
            </tr>";

        foreach($trs as $tr){
            if($tr != ''){
                $link = array();
                $pattern = "/(<a href='.*?'>.*?<\/a>)|((http:\/\/|ftp:\/\/|https:\/\/|\/))?[\w-]+(\.[\w-]+)+([\w.,@?^=%&amp;:\/~+#-]*[\w@?^=%&amp;\/~+#-])?/i";
                if(preg_match($pattern, $tr, $link)){
                    $tr = preg_replace($pattern, '', $tr, 1);
                }

                $link = (isset($link[0]) && $link[0] != '') ? $link[0] : 'ссылка не найдена';
                $td = htmlspecialchars($tr);
                $xls .= "<tr>";
                $xls .= "<td valign='top' align='left' width='100'>$link</td>";
                $xls .= "<td valign='top' align='left' width='700'>$td</td>";
                $xls .= "</tr>";
            }
        }

        $xls .= "</table></body></html>";

        header("Content-Type: application/excel; charset=utf-8");
        header('Content-Disposition: attachment; filename="xls.xls"');
        header("Content-Length: " . strlen($xls));

        echo $xls;
        exit();
    }

    public function  sendDownloadFile($container=true, $generate_urls=false){
        $this->getHtml(false, $generate_urls);
        $filename = 'html.html';

        header("Content-Type: text/plain; charset=utf-8");
        header('Content-Disposition: attachment; filename="'.$filename.'"');
        header("Content-Length: " . strlen($this->output));

        if($container){
            echo "<!DOCTYPE html>
            <head>
                <meta charset='UTF-8' />
                <title></title>
            </head>
            <body>";
        }

        echo $this->output;

        if($container) {
            echo "</body></html>";
        }

        exit;
    }

    private function getElements(&$innerElements='', $node, $nodeValue, &$tag, $tagElement=array(), $generates_urls=false)
    {
        foreach ($node->childNodes as $element) {
            if($element->hasChildNodes()){
                $elementAlreadyCreate = false;
                //haveChild
                if($element->nodeName == 'w:pPr') {
                    $thisAttr = $this->getAttr($element, 'val');
                    if(is_numeric($thisAttr)) {
                        $hNumb = rtrim($thisAttr, "0");

                        $tag['open'] = '<h'.$hNumb.'>';
                        $tag['close'] = '</h'.$hNumb.'>';
                    }

                    if($this->hasElement($element, 'numPr'))
                    {
                        $rid = $element->getElementsByTagName('numId')->item(0)->attributes->item(0)->value;

                        $type = $this->extractNumberingXML($rid);

                        if(!isset($tag['open'])) {
                            $tag['open'] = '';
                        }
                        if(!isset($tag['close'])) {
                            $tag['close'] = '';
                        }

                        $tag['open'] = '<li data-type="'.$type.'">'.$tag['open'];
                        $tag['close'] = $tag['close'].'</li>';
                    }
                }
                elseif($element->nodeName == 'w:rPr') {
                    $tagElement = [];

                    $this->generateTags($element, $tagElement);

                    if($this->hasElement($element, 'rStyle')){
                        //Fucking tags
                        $thisAttr = $element->getElementsByTagName('rStyle')->item(0)->attributes->getNamedItem('val')->nodeValue;
                        if(is_numeric($thisAttr) && !$this->hasElement($element->parentNode->parentNode, 'pStyle')){
                            $hNumb = rtrim($thisAttr, "0");

                            $tag['open'] = '<h'.$hNumb.'>';
                            $tag['close'] = '</h'.$hNumb.'>';
                        }
                    }
                }
                elseif($element->nodeName == 'w:hyperlink') {
                    //Links
                    $rid = $element->attributes->getNamedItem('id')->nodeValue;
                    $rels = $this->extractRelXML();
                    $path = $rels[$rid][0];
                    $target = $rels[$rid][3];

                    $tagLink = '<a>';
                    if (strtolower($target) == "external") {
                        $tagLink = "<a href='" . $path . "'>";
                    } elseif (isset($element->attributes->getNamedItem('anchor')->nodeValue)) {
                        $tagLink = "<a href='#" . $element->attributes->getNamedItem('anchor')->nodeValue . "'>";
                    }

                    $tagElementThis = array();
                    $this->generateTags($element->getElementsByTagName('rPr')->item(0), $tagElementThis);

                    $innerElements .= $tagLink.$tagElementThis['open'].$element->nodeValue.$tagElementThis['close'].'</a>';
                    $elementAlreadyCreate = true;
                }
                elseif($element->nodeName == 'w:t') {
                    //Simple text
                    if(!isset($tagElement['open'])) {
                        $tagElement['open'] = '';
                    }
                    if(!isset($tagElement['close'])) {
                        $tagElement['close'] = '';
                    }

                    $content = ($generates_urls) ? $this->generateUrlsToTags($element->nodeValue) : $element->nodeValue;
                    $innerElements .= $tagElement['open'].$content.$tagElement['close'];

                }

                if(!$elementAlreadyCreate) $this->getElements($innerElements, $element, $nodeValue, $tag, $tagElement, $generates_urls);

            }
        }
        return;
    }

    private function generateUrlsToTags($content){
        $pattern = '/(^http:\/\/.+)/i';
        if(preg_match($pattern, $content)){
            $replacement = '<a href="$1">$1</a>';
            $content = preg_replace($pattern, $replacement, $content);
        }

        return $content;
    }

    private function generateTags($node, &$tag=array(), $table=false) {
        if(!isset($tag['open'])) {
            $tag['open'] = '';
        }
        if(!isset($tag['close'])) {
            $tag['close'] = '';
        }

        if ($node->childNodes) {
            if ($table) {
                if ($this->hasElement($node, 'b')) {
                    $tag['open'] .= '<b>';
                    $tag['close'] .= '</b>';
                } elseif ($this->hasElement($node, 'u')) {
                    $tag['open'] .= '<u>';
                    $tag['close'] .= '</u>';
                } elseif ($this->hasElement($node, 'i')) {
                    $tag['open'] .= '<i>';
                    $tag['close'] .= '</i>';
                } elseif ($this->hasElement($node, 'strike')) {
                    $tag['open'] .= '<strike>';
                    $tag['close'] .= '</strike>';
                }
            }

            foreach ($node->childNodes as $element) {
                switch($element->nodeName){
                    case 'w:b': {
                        //Bold
                        $tag['open'] .= '<b>';
                        $tag['close'] .= '</b>';
                        break;
                    }
                    case 'w:u': {
                        //Underline
                        $tag['open'] .= '<u>';
                        $tag['close'] .= '</u>';
                        break;
                    }
                    case 'w:i': {
                        //Italic
                        $tag['open'] .= '<em>';
                        $tag['close'] .= '</em>';
                        break;
                    }
                    case 'w:strike': {
                        //Italic
                        $tag['open'] .= '<strike>';
                        $tag['close'] .= '</strike>';
                        break;
                    }

                }

            }
        }
        return;
    }

    private function getAttr($element, $val){
        return (isset($element->firstChild->attributes->getNamedItem($val)->nodeValue)) ? $element->firstChild->attributes->getNamedItem($val)->nodeValue : '';
    }

    private function hasElement($node, $has){
        if($node->getElementsByTagName($has)->length > 0)
            return true;
        else
            return false;
    }

    private function reformatOutput($output){
        //remove data attributes
        $output = preg_replace('/(data-type="ul")|(data-type="ol")/', '', $output);

        return $output;
    }
}
