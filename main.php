<?php

    require 'vendor/autoload.php';
    require 'src/class.application.base.php';
    require 'src/class.application.php';

    $pathHtml = $argv[1];
    $pathDocx = $argv[2];

    $wordInfo   = Array(

    );

    $wordConfig = Array(
        "zoom"                  => 100,   // Int.
        "mirrorMargins"         => FALSE, // Bool.
        "decimalSymbol"         => ",",   //String (, | .).
        "themeFontLang"         => 'ES_AR', // String || NULL.

        // Corrector gramatical y ortografico.
        "hideSpellingErrors"    => TRUE,  // Bool.
        "hideGrammaticalErrors" => TRUE,  // Bool.
        "proofState"            => 'CLEAN', // String (DIRTY | CLEAN).

        // Seguimiento de cambios.
        "trackRevisions"        => TRUE,  // Bool.
        "doNotTrackMoves"       => TRUE,  // Bool.
        "doNotTrackFormatting"  => TRUE,  // Bool.


        "revisionView"          => TRUE,  // Bool.
        "documentProtection"    => Array(TRUE,'READ_ONLY','password'),
        
        "updateFields"          => TRUE,  // Bool.
        
        // Separación.
        "autoHyphenation"       => TRUE,  // Bool.
        "consecutiveHyphenLimit"=> 0,     // Int.
        "hyphenationZone"       => \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1), // Object.
        "doNotHyphenateCaps"    => TRUE   // Bool.
     );

    $application = new Application($pathHtml,$pathDocx);
    $application->createDocx($wordInfo,$wordConfig);

?>