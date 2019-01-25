<?php
    class Application extends ApplicationBase{

        private $html;
        private $XMLHead;
        private $XMLBody;

        private function loadHTML(){
            $this->html = file_get_contents($this->htmlPath);
        }

        private function normalizeHTML(){
            $this->html = str_replace(
                $this->charsHTM,
                $this->charsDOC,
                $this->html
            );

            $this->html = preg_replace(
                $this->stringHT5A,
                '<div',
                $this->html
            );

            $this->html = preg_replace(
                $this->stringHT5B,
                '</div>',
                $this->html
            );
        }

        private function transformXML(){
            $this->HTML->loadHTML(
                $this->html
            );

            /*$style = '';
            $styles = $this->HTML->getElementsByTagName('style');
            foreach($styles as $s) if($s->tagName=='style') $style .= $s->textContent;
            $style = str_replace(["\n","\t","\b"],"",$style);
            echo($style);*/
            //print_r(explode($style,'.'));
            //$styleN = explode(,'.');
            //print_r($styleN);

            /*$this->XMLBody = $this->HTML->getElementsByTagName("body");

            echo "Header:\n";
            //print_r($this->XMLHead);
            print_r($this->XMLHead[0]);
            //foreach($this->XMLHead as $head) print_r($head);
            echo "Body:\n";
            print_r($this->XMLBody);*/

            $body = $this->HTML->getElementsByTagName("body");
            print_r($body->childNodes);
        }

        public function __construct($pathHtml,$pathDocx){
            parent::__construct($pathHtml,$pathDocx);
        }
        public function createDocx(){
            if($this->boolPathHtml==TRUE && $this->boolPathDocx==true){

                $this->loadHTML();
                $this->normalizeHTML();
                $this->transformXML();

                //$this->setHtml();

                /*$section = $this->Word->addSection();
                //$section->addText("Hola Mundo!");
                \PhpOffice\PhpWord\Shared\Html::addHtml($section,$this->html,true,true);

                $docxPath = basename($this->htmlPath,".html");
                $docxPath = $this->docxPath . "/" . $docxPath . ".docx";
                $this->Log->info("Crear archivo en: " . $docxPath);

                $WordWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->Word,"Word2007");
                $WordWriter->save($docxPath);*/
                
            } else $this->Log->error("Se ha detenido la construccion del documento.");
        }
        public function __destruct(){
            parent::__destruct();
        }
    }