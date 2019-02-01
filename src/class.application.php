<?php

    use \PhpOffice\PhpWord\Style\Language;
    use \PhpOffice\PhpWord\ComplexType\ProofState;
    use \PhpOffice\PhpWord\SimpleType\DocProtect;
    use \PhpOffice\PhpWord\Shared\Converter;

    class Application extends ApplicationBase {

        private $fontName;
        private $fontSize;
        private $zoom;

        public function __construct($pathHtml,$pathDocx)
        {

            parent::__construct($pathHtml,$pathDocx);

        }

        public function createDocx($default,$info,$config)
        {

            $this->Word->setDefaultFontName($default['fontName']);
            $this->Word->setDefaultFontSize($default['fontSize']);

            $properties = $this->Word->getDocInfo();
            $properties->setCreator($info['creator']);
            $properties->setCompany($info['company']);
            $properties->setTitle($info['title']);
            $properties->setDescription($info['description']);
            $properties->setCategory($info['category']);
            $properties->setLastModifiedBy($info['modifiedBy']);
            $properties->setCreated(mktime($info['created'][0], $info['created'][1], $info['created'][2], $info['created'][3], $info['created'][4], $info['created'][5]));
            $properties->setModified(mktime($info['modified'][0], $info['modified'][1], $info['modified'][2], $info['modified'][3], $info['modified'][4], $info['modified'][5]));
            $properties->setSubject($info['subject']);
            $properties->setKeywords($info['keywords']);

            if(is_int($config["zoom"])==TRUE) $this->Word->getSettings()->setZoom($config["zoom"]);
            if($config["mirrorMargins"]==TRUE)$this->Word->getSettings()->setMirrorMargins(TRUE);
            if($config["decimalSymbol"]==".") $this->Word->getSettings()->setDecimalSymbol(".");
            if($config["decimalSymbol"]==",") $this->Word->getSettings()->setDecimalSymbol(",");
            //if($config["revisionView"]==TRUE) $this->Word->getSettings()->setRevisionView(TRUE);
            if($config["updateFields"]==TRUE) $this->Word->getSettings()->setUpdateFields(TRUE);
            if(!is_null($config["themeFontLang"]))
            {
                $lang = new Language();
                $lang->setLangId(Language::ES_ES);
                $this->Word->getSettings()->setThemeFontLang($lang);
            }

            // Corrector gramatical y ortografico.
            if($config["hideSpellingErrors"]==TRUE)    $this->Word->getSettings()->setHideSpellingErrors(TRUE);
            if($config["hideGrammaticalErrors"]==TRUE) $this->Word->getSettings()->setHideGrammaticalErrors(TRUE);
            if($config["proofState"]=='CLEAN')
            {
                $state = new ProofState();
                $state->setGrammar(ProofState::CLEAN);
                $state->setSpelling(ProofState::CLEAN);
                $this->Word->getSettings()->setProofState($state);
            }
            if($config["proofState"]=='DIRTY')
            {
                $state = new ProofState();
                $state->setGrammar(ProofState::DIRTY);
                $state->setSpelling(ProofState::DIRTY);
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
                $documentProtection->setPassword($config["documentProtection"][2]);
                if($config["documentProtection"][1]=='NONE') $documentProtection->setEditing(DocProtect::NONE);
                if($config["documentProtection"][1]=='READ_ONLY') $documentProtection->setEditing(DocProtect::READ_ONLY);
                if($config["documentProtection"][1]=='COMMENTS') $documentProtection->setEditing(DocProtect::COMMENTS);
                if($config["documentProtection"][1]=='TRACKED_CHANGES') $documentProtection->setEditing(DocProtect::TRACKED_CHANGES);
                if($config["documentProtection"][1]=='FORMS') $documentProtection->setEditing(DocProtect::FORMS);
            }
            
             // Separación.
            if($config["autoHyphenation"]==TRUE) $this->Word->getSettings()->setAutoHyphenation(TRUE);
            if(is_int($config["consecutiveHyphenLimit"])==TRUE) $this->Word->getSettings()->setConsecutiveHyphenLimit($config["consecutiveHyphenLimit"]);
            if($config["doNotHyphenateCaps"]==TRUE) $this->Word->getSettings()->setDoNotHyphenateCaps(TRUE);
            if($config["hyphenationZone"]>=1)
            {
                $this->Word->getSettings()->setHyphenationZone(
                    Converter::cmToTwip($config["hyphenationZone"])
                );
            }

            $this->loadHTML();

            $this->normalizeHTML();

            $this->parseCssStyleToWordStyle();

            $this->parteHtmlBodyToSectionWord();

            $this->saveDocument();

        }
        public function __destruct(){
            parent::__destruct();
        }
    }