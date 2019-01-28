<?php

    use \Katzgrau\KLogger\Logger;
    use \PhpOffice\PhpWord\PhpWord;
    use \Symfony\Component\DomCrawler\Crawler;
    use \Symfony\Component\CssSelector\CssSelectorConverter;

    class ApplicationBase{

        protected $Log;
        protected $Dom;
        protected $Html;
        protected $Word;
        protected $Selector;
        protected $XMLDom;

        protected $charsHTM = ['<br>','<BR>','<br/>','<BR/>','&nbsp;'];
        protected $charsDOC = ['\n','\n','\n','\n','\ '];

        protected $stringHT5A = '/(<header)|(<nav)|(<section)|(<article)|(<aside)|(<footer)/i';
        protected $stringHT5B = '/(<\/header>)|(<\/nav>)|(<\/section>)|(<\/article>)|(<\/aside>)|(<\/footer>)/i';

        protected $htmlPath;
        protected $docxPath;
        protected $boolPathHtml;
        protected $boolPathDocx;

        protected function loadHTML(){
            $this->Html = file_get_contents($this->htmlPath);
            $this->Log->info("Se ha cargado el contenido del archivo html.");
        }

        protected function normalizeHTML(){

            $this->Html = str_replace(
                $this->charsHTM,
                $this->charsDOC,
                $this->Html
            );

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

            $this->Log->info("Se ha normatlizado el contenido html.");

        }

        protected function parseHTMLToDOMElements(){

            $this->Dom = new Crawler(
                $this->Html
            );
            $this->Log->info("Clase Symfony/Crawler instanciada.");

            $this->Selector = new CssSelectorConverter();
            $this->Log->info("Clase Symfony/CSS-Selector instanciada.");

        }

        protected function parseCssStyleToWordStyle()
        {
            $style = '';
            $nodesStyles = $this->Dom->filter("style");
            foreach($nodesStyles as $n) if($n->nodeName=="style") $style .= $n->nodeValue;

            $arrayStyle = explode('.',$style);
            foreach($arrayStyle as $s){

                $iniAnchor = strpos($s,'{');
                $endAnchor = strpos($s,'}');

                $styleName = substr(
                    $s,
                    0,
                    $iniAnchor
                );

                $fontName = null;
                $fontSize = null;
                $color = null;
                $bold = null;

                $this->Word->addFontStyle(
                    $styleName,
                    array(
                        "name" => $fontName,
                        "size" => $fontSize,
                        "color" => $color,
                        "bold" => $bold

                    )
                );

            }
        }

        public function __construct($pathHtml,$pathDocx){

            $this->Log = new Logger(__DIR__.'/../logs');
            $this->Log->info('Clase KLogger instanciada.');

            $this->Word = new PhpWord();
            $this->Log->info('Clase PhpWord instanciada.');

            $htmlPath = array();
            $docxPath = array();
            $pattern = "/[a-z0-9\\\:\/\-\_\.\ ]+/i";
            
            $this->boolPathHtml = preg_match($pattern,$pathHtml,$htmlPath);
            $this->boolPathDocx = preg_match($pattern,$pathDocx,$docxPath);

            if($this->boolPathDocx==TRUE && $this->boolPathDocx==TRUE)
            {

                if(file_exists($htmlPath[0])==TRUE && is_dir($docxPath[0])==TRUE)
                {
                    $this->htmlPath = $htmlPath[0];
                    $this->docxPath = $docxPath[0];

                    $this->loadHTML();
                    $this->normalizeHTML();
                    $this->parseHTMLToDomElements();
                    $this->parseCssStyleToWordStyle();

                }
                else $this->Log->error("Los parámetros enviados son incorrectos.");
                
            } else $this->Log->error("Falla algún parámetro.");
        }

        public function __destruct()
        {

        }
    }