<?php
declare(strict_types=1);
final class BitrixArticleXmlExporter {
    private const PROPS=['ARTICLE_TYPE'=>847,'PRIMARY_QUERY'=>848,'SECONDARY_QUERIES'=>849,'SEARCH_INTENT'=>850,'SHORT_ANSWER'=>851,'REGION'=>852,'AUTHOR'=>853,'MEDICAL_REVIEWER'=>854,'MEDICAL_REVIEWED_AT'=>855,'CONTENT_UPDATED_AT'=>856,'SOURCES'=>857,'RELATED_ARTICLES'=>858,'ARTICLE_TEMPLATE'=>864];
    public function build(array $d): array {
        $warnings=[];$this->validate($d);$code=$this->safeCode((string)$d['result_code']);$date=date('Ymd');
        $doc=new DOMDocument('1.0','UTF-8');$doc->formatOutput=true;
        $root=$doc->createElement('КоммерческаяИнформация');$root->setAttribute('ВерсияСхемы','2.021');$doc->appendChild($root);
        $classifier=$root->appendChild($doc->createElement('Классификатор'));$classifier->appendChild($doc->createElement('Ид','81'));$classifier->appendChild($doc->createElement('Наименование','Медицинские статьи'));
        $props=$classifier->appendChild($doc->createElement('Свойства'));foreach(self::PROPS as $codeProp=>$id){$p=$props->appendChild($doc->createElement('Свойство'));$p->appendChild($doc->createElement('Ид',(string)$id));$p->appendChild($doc->createElement('Код',$codeProp));}
        $catalog=$root->appendChild($doc->createElement('Каталог'));$catalog->appendChild($doc->createElement('Ид','81'));$catalog->appendChild($doc->createElement('Код','medical_articles_v2'));$catalog->appendChild($doc->createElement('Наименование','Медицинские статьи'));
        $products=$catalog->appendChild($doc->createElement('Товары'));$t=$products->appendChild($doc->createElement('Товар'));
        $t->appendChild($doc->createElement('Ид','medical-article-'.$code.'-'.$date));$t->appendChild($doc->createElement('Наименование',(string)$d['result_name']));
        $this->prop($doc,$t,'CML2_ACTIVE','false');$this->prop($doc,$t,'CML2_SORT','500');$this->prop($doc,$t,'CML2_CODE',$code);
        $this->textProp($doc,$t,'CML2_PREVIEW_TEXT',(string)($d['result_preview']??''),'text');$this->textProp($doc,$t,'CML2_DETAIL_TEXT',$this->sanitizeHtml((string)$d['result_detail_html']),'html');
        if(!empty($d['article_section_id'])){$groups=$t->appendChild($doc->createElement('Группы'));$groups->appendChild($doc->createElement('Ид',(string)$d['article_section_id']));}
        $values=$t->appendChild($doc->createElement('ЗначенияСвойств'));$filled=0;
        $map=[847=>'article_type_xml_id',848=>'primary_query',850=>'search_intent_xml_id',851=>'result_short_answer',852=>'region_xml_id',864=>'article_template_xml_id',853=>'author_id',854=>'medical_reviewer_id',855=>'medical_reviewed_at',856=>'content_updated_at'];
        foreach($map as $id=>$key){$val=$d[$key]??($key==='article_template_xml_id'?'default':'');if($val!==''){$this->xmlValue($doc,$values,$id,$this->dateRus((string)$val));$filled++;}}
        foreach($this->list($d['secondary_queries']??[]) as $v){$this->xmlValue($doc,$values,849,$v);$filled++;}
        foreach($this->list($d['sources']??[]) as $v){$this->xmlValue($doc,$values,857,$v);$filled++;}
        foreach($this->related($d['related_articles']??[],$warnings) as $v){$this->xmlValue($doc,$values,858,$v);$filled++;}
        return ['xml'=>$doc->saveXML(),'filename'=>'medical-article-'.$code.'-'.$date.'.xml','filled'=>$filled,'warnings'=>$warnings];
    }
    private function validate(array $d): void {foreach(['result_code','result_name','result_detail_html'] as $f)if(trim((string)($d[$f]??''))==='')throw new InvalidArgumentException('Required field: '.$f);if(!$this->isUtf8($d))throw new InvalidArgumentException('Only UTF-8 payload is supported');if(preg_match('~<script\b|<iframe\b(?![^>]+src=["\']https://(?:www\.)?(?:youtube\.com|rutube\.ru)/)~iu',(string)$d['result_detail_html']))throw new InvalidArgumentException('Dangerous HTML is not allowed');}
    private function isUtf8($v): bool {if(is_array($v))foreach($v as $x){if(!$this->isUtf8($x))return false;}return !is_string($v)||mb_check_encoding($v,'UTF-8');}
    private function safeCode(string $s): string {$s=preg_replace('~[^a-z0-9_-]+~i','-',trim($s));return trim($s,'-')?:'article';}
    private function sanitizeHtml(string $s): string {$s=preg_replace('~<\/?(script|style|iframe|form)[^>]*>.*?<\/\1>~isu','',$s);return str_replace(']]>',']]]]><![CDATA[>',$s);}
    private function dateRus(string $s): string {if(preg_match('~^\d{4}-\d{2}-\d{2}$~',$s))return date('d.m.Y',strtotime($s));return $s;}
    private function list($v): array {if(is_string($v))$v=preg_split('~\r?\n|,~',$v);return array_values(array_filter(array_map('trim',(array)$v),fn($x)=>$x!==''));}
    private function related($v,array &$warnings): array {$out=[];foreach($this->list($v) as $item){if(str_contains($item,':')){[$ib,$id]=explode(':',$item,2);if(trim($ib)==='81'){$warnings[]='Связанная статья инфоблока 81 исключена из свойства 858: '.$id;continue;}$out[]=trim($id);}else{$out[]=$item;}}return $out;}
    private function prop(DOMDocument $d,DOMElement $p,string $n,string $v): void {$p->appendChild($d->createElement($n,$v));}
    private function textProp(DOMDocument $d,DOMElement $p,string $n,string $v,string $type): void {$e=$p->appendChild($d->createElement($n));$e->setAttribute('Тип',$type);$e->appendChild($d->createCDATASection($v));}
    private function xmlValue(DOMDocument $d,DOMElement $values,int $id,string $value): void {$z=$values->appendChild($d->createElement('ЗначенияСвойства'));$z->appendChild($d->createElement('Ид',(string)$id));$z->appendChild($d->createElement('Значение',$value));}
}
