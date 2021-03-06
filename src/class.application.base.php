<?php

    use \Katzgrau\KLogger\Logger;
    
    use \Symfony\Component\DomCrawler\Crawler;
    use \Symfony\Component\CssSelector\CssSelectorConverter;

    use \PhpOffice\PhpWord\PhpWord;
    use \PhpOffice\PhpWord\Shared\Converter;
    use \PhpOffice\PhpWord\SimpleType\Jc;
    use \PhpOffice\PhpWord\SimpleType\JcTable;
    use \PhpOffice\PhpWord\Style\Font;
    use \PhpOffice\PhpWord\SimpleType\TextAlignment;
    use \PhpOffice\PhpWord\IOFactory;

    class ApplicationBase{

        protected $Log;
        protected $Dom;
        protected $DomXPath;
        protected $Html;
        protected $Word;
        protected $Section;
        protected $WordStyles = Array();
        protected $XMLDom;

        protected $stringHT5A = '/(<header)|(<nav)|(<section)|(<article)|(<aside)|(<footer)/i';
        protected $stringHT5B = '/(<\/header>)|(<\/nav>)|(<\/section>)|(<\/article>)|(<\/aside>)|(<\/footer>)/i';

        protected $htmlPath;
        protected $docxPath;
        protected $boolPathHtml;
        protected $boolPathDocx;

        protected function loadHTML()
        {
            $this->Html = file_get_contents($this->htmlPath);
            $this->Log->info("Se ha cargado el contenido del archivo html.");
        }

        protected function normalizeHTML()
        {

            $this->Html = preg_replace(
                $this->stringHT5A,
                '<div',
                $this->Html
            );

            $this->Html = preg_replace(
                $this->stringHT5B,
                '</div>',
                $this->Html
            );

            $this->Html = str_replace(array("\n", "\r"), '', $this->Html);
            $this->Html = str_replace(array('&lt;', '&gt;', '&amp;'), array('_lt_', '_gt_', '_amp_'), $this->Html);
            $this->Html = html_entity_decode($this->Html, ENT_QUOTES, 'UTF-8');
            $this->Html = str_replace('&', '&amp;', $this->Html);
            $this->Html = str_replace(array('_lt_', '_gt_', '_amp_'), array('&lt;', '&gt;', '&amp;'), $this->Html);

            // Implementar luego.
            /*if (false === $fullHTML) {
                $html = '<body>' . $html . '</body>';
            }*/
            $this->Log->info("Se ha normatlizado el contenido html.");

        }

        protected function parseCssStyleToWordStyle()
        {
            $style = NULL;

            libxml_disable_entity_loader(true);
            $this->Dom =new DOMDocument();
            $this->Dom->preserveWhiteSpace = false;
            $this->Dom->loadHTML($this->Html);

            $this->DomXPath = new DOMXpath($this->Dom);

            $list = $this->DomXPath->query("/html/head/style");

            foreach($list as $n) if($n->nodeName=="style") $style .= $n->nodeValue;
            $style = trim($style);

            $arrayStyle = explode('.',$style);
            
            foreach($arrayStyle as $s)
            {
                if(!$s=="")
                {
                    $iniAnchor = strpos($s,'{');
                    $endAnchor = strpos($s,'}');
    
                    $styleName = substr(
                        trim($s),
                        0,
                        $iniAnchor
                    );

                    $styleContent = substr(
                        trim($s),
                        $iniAnchor,
                        $endAnchor
                    );

                    $styleContent = str_replace(
                        ['{','}'],
                        '',
                        $styleContent
                    );

                    $styleLines = explode(
                        ';',
                        trim($styleContent)
                    );

                    $styleWord = Array();

                    foreach($styleLines as $l){

                        $l = trim($l);

                        $css = explode(
                            ":",
                            $l
                        );

                        if(count($css)==2){
                            $title = strtolower($css[0]);
                            $text = $css[1];

                            if($title=='font-family') $styleWord['name'] = substr($text,0,strpos($text,','));      //{ $title='name';  $text=substr($text,0,strpos($text,','));     }
                            if($title=='font-size')   $styleWord['size'] = Converter::pixelToTwip(intval($text)); //{ $title='size';  $text=Converter::pixelToTwip(intval($text)); }
                            if($title=='font-weight') $styleWord['bold'] = TRUE; //{ $title='bold';  $text=TRUE; }
                            if($title=='text-align')  {
                                if($text=='justify') $text = Jc::JUSTIFY;
                                if($text=='left')    $text = Jc::LEFT;
                                if($text=='center')  $text = Jc::CENTER;
                                if($text=='right')   $text = Jc::RIGHT;
                                $styleWord['alignment'] = $text;
                            }
                            if($title=='vertical-align'){
                                if($text=='top')      $text = TextAlignment::START;
                                if($text=='middle')   $text = TextAlignment::CENTER;
                                if($text=='baseline') $text = TextAlignment::BASELINE;
                                $styleWord['textAlignment'] = $text;
                            }
                            if($title=='text-transform'){}

                            

                            /*
                            array_push(
                                $styleWord,
                                Array(
                                    'name' => $title,
                                    'value' => $text
                                )
                            );
                            Font
                            ====
                                allCaps. All caps, true or false.
                                bgColor. Font background color, e.g. FF0000.

                                color. Font color, e.g. FF0000.
                                doubleStrikethrough. Double strikethrough, true or false.
                                fgColor. Font highlight color, e.g. yellow, green, blue.
                                See \PhpOffice\PhpWord\Style\Font::FGCOLOR_... constants for more values
                                hint. Font content type, default, eastAsia, or cs.
                                italic. Italic, true or false.
                                rtl. Right to Left language, true or false.
                                smallCaps. Small caps, true or false.
                                strikethrough. Strikethrough, true or false.
                                subScript. Subscript, true or false.
                                superScript. Superscript, true or false.
                                underline. Underline, single, dash, dotted, etc.
                                    See \PhpOffice\PhpWord\Style\Font::UNDERLINE_... constants for more values
                                lang. Language, either a language code like en-US, fr-BE, etc. or an object (or as an array) if you need to set eastAsian or bidirectional languages
                                    See \PhpOffice\PhpWord\Style\Language class for some language codes.
                                position. The text position, raised or lowered, in half points


                            Paragraph
                            =========
                                    basedOn. Parent style.
                                    hanging. Hanging in twip.
                                    indent. Indent in twip.
                                    keepLines. Keep all lines on one page, true or false.
                                    keepNext. Keep paragraph with next paragraph, true or false.
                                    lineHeight. Text line height, e.g. 1.0, 1.5, etc.
                                    next. Style for next paragraph.
                                    pageBreakBefore. Start paragraph on next page, true or false.
                                    spaceBefore. Space before paragraph in twip.
                                    spaceAfter. Space after paragraph in twip.
                                    spacing. Space between lines.
                                    spacingLineRule. Line Spacing Rule. auto, exact, atLeast
                                    suppressAutoHyphens. Hyphenation for paragraph, true or false.
                                    tabs. Set of custom tab stops.
                                    widowControl. Allow first/last line to display on a separate page, true or false.
                                    contextualSpacing. Ignore Spacing Above and Below When Using Identical Styles, true or false.
                                    bidi. Right to Left Paragraph Layout, true or false.
                                    shading. Paragraph Shading.



                            */
                        }

                    }

                    $this->WordStyle[$styleName] = $styleWord;
                }

            }
        }

        protected function parteHtmlBodyToSectionWord()
        {

            $this->Section = $this->Word->addSection();
            Html::addHtml($this->Section, $this->Dom, $this->DomXPath, $this->WordStyle);

        }

        protected function saveDocument()
        {

            $WordWriter = IOFactory::createWriter($this->Word,"Word2007");
            $WordWriter->save($this->docxPath);

        }

        public function __construct($pathHtml,$pathDocx)
        {

            $this->Log = new Logger(__DIR__.'/../logs');
            $this->Log->info('Clase KLogger instanciada.');

            $this->Word = new PhpWord();
            $this->Log->info('Clase PhpWord instanciada.');

            $this->Dom = new DOMDocument();
            $this->Log->info("Clase DOMDocument instanciada.");

            $htmlPath = array();
            $docxPath = array();
            $pattern = "/[a-z0-9\\\:\/\-\_\.\ ]+/i";
            
            $this->boolPathHtml = preg_match($pattern,$pathHtml,$htmlPath);
            $this->boolPathDocx = preg_match($pattern,$pathDocx,$docxPath);

            if($this->boolPathDocx==TRUE && $this->boolPathDocx==TRUE)
            {

                if(file_exists($htmlPath[0])==TRUE && is_dir($docxPath[0])==TRUE)
                {
                    $docx  = basename($htmlPath[0],".html");
                    $docx .= ".docx";

                    $this->htmlPath = $htmlPath[0];
                    $this->docxPath = $docxPath[0] . "/" .$docx;
                    
                }
                else $this->Log->error("Los parámetros enviados son incorrectos.");
                
            } else $this->Log->error("Falla algún parámetro.");
        }

        public function __destruct()
        {
        }
    }