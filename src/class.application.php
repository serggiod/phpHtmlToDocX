<?php
    class Application extends ApplicationBase{

        private $fontName;
        private $fontSize;
        private $zoom;

        public function __construct($pathHtml,$pathDocx)
        {

            parent::__construct($pathHtml,$pathDocx);

        }

        public function createDocx($info,$config)
        {
            if(is_int($config["zoom"])==TRUE) $this->Word->getSettings()->setZoom($config["zoom"]);
            if($config["mirrorMargins"]==TRUE)$this->Word->getSettings()->setMirrorMargins(TRUE);
            if($config["decimalSymbol"]==".") $this->Word->getSettings()->setDecimalSymbol(".");
            if($config["decimalSymbol"]==",") $this->Word->getSettings()->setDecimalSymbol(",");
            if($config["revisionView"]==TRUE) $this->Word->getSettings()->setRevisionView(TRUE);
            if($config["updateFields"]==TRUE) $this->Word->getSettings()->setUpdateFields(TRUE);
            if(!is_null($config["themeFontLang"]))
            {
                $lang = new Language();
                $lang->setLangId(Language::$config["themeFontLang"]);
                $this->Word->getSettings()->setThemeFontLang($lang);
            }

            // Corrector gramatical y ortografico.
            if($config["hideSpellingErrors"]==TRUE)    $this->Word->getSettings()->setHideSpellingErrors(TRUE);
            if($config["hideGrammaticalErrors"]==TRUE) $this->Word->getSettings()->setHideGrammaticalErrors(TRUE);
            if($config["proofState"]=='CLEAN'||$config["proofState"]=='DIRTY')
            {
                $state = new ProofState();
                $state->setGrammar(ProofState::$config["proofState"]);
                $state->setSpelling(ProofState::$config["proofState"]);
                $this->Word->getSettings()->setProofState($state);
            } 
            
            // Seguimiento de cambios.
            if($config["trackRevisions"      ]==TRUE) $this->Word->getSettings()->setTrackRevisions(TRUE);
            if($config["doNotTrackMoves"     ]==TRUE) $this->Word->getSettings()->setDoNotTrackMoves(TRUE);
            if($config["doNotTrackFormatting"]==TRUE) $this->Word->getSettings()->setDoNotTrackFormatting(TRUE);
            
            // Protección por password.
            if($config["documentProtection"][0]==TRUE)
            {
                $documentProtection = $this->Word->getSettings()->getDocumentProtection();
                $documentProtection->setEditing(DocProtect::$config["documentProtection"][1]);
                $documentProtection->setPassword($config["documentProtection"][2]);
            }
            
            
            // Separación.
            if($config["autoHyphenation"]==TRUE) $this->Word->getSettiongs()->setAutoHyphenation(TRUE);
            if(is_int($config["consecutiveHyphenLimit"])==TRUE) $this->Word->getSettiongs()->setConsecutiveHyphenLimit($config["consecutiveHyphenLimit"]);
            if($config["doNotHyphenateCaps"]==TRUE) $this->Word->getSettiongs()->setDoNotHyphenateCaps(TRUE);
            if($config["hyphenationZone"]==TRUE)
            {
                $this->Word->getSettiongs()->setHyphenationZone(
                    \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1)
                );
            }

            //$this->Word->setDefaultFontName($this->fontName);
            //$this->Word->setDefaultFontSize($this->fontSize);

        }
        public function __destruct(){
            parent::__destruct();
        }
    }