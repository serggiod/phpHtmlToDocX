<?php
    class Application extends ApplicationBase{
        private $html;

        private function setHtml(){
            $this->html = file_get_contents($this->htmlPath);
        }

        public function __construct($pathHtml,$pathDocx){
            parent::__construct($pathHtml,$pathDocx);
        }
        public function createDocx(){
            if($this->boolPathHtml==TRUE && $this->boolPathDocx==true){
                
                $this->setHtml();

                $section = $this->Word->addSection();
                //$section->addText("Hola Mundo!");
                \PhpOffice\PhpWord\Shared\Html::addHtml($section,$this->html,false,false);

                $docxPath = basename($this->htmlPath,".html");
                $docxPath = $this->docxPath . "/" . $docxPath . ".docx";
                $this->Log->info("Crear archivo en: " . $docxPath);

                $WordWriter = \PhpOffice\PhpWord\IOFactory::createWriter($this->Word,"Word2007");
                $WordWriter->save($docxPath);
                
            } else $this->Log->error("Se ha detenido la construccion del documento.");
        }
        public function __destruct(){
            parent::__destruct();
        }
    }