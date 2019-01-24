<?php

    class ApplicationBase{

        protected $Log;
        protected $Word;

        protected $htmlPath;
        protected $docxPath;
        protected $boolPathHtml;
        protected $boolPathDocx;

        public function __construct($pathHtml,$pathDocx){

            $this->Log = new Katzgrau\KLogger\Logger(__DIR__.'/../logs');
            $this->Log->info('Clase KLogger Instanciada.');

            $this->Word = new \PhpOffice\PhpWord\PhpWord();
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
                    $this->Log->info("Parámetros correctos.");
                }
                else $this->Log->error("Los parámetros enviados son incorrectos.");
                
            } else $this->Log->error("Falla algún parámetro.");
        }

        public function __destruct()
        {

        }
    }