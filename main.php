<?php

    require 'vendor/autoload.php';
    require 'src/class.application.base.php';
    require 'src/class.application.php';

    $pathHtml = $argv[1];
    $pathDocx = $argv[2];

    // Valores por defecto.
    $default = Array(
        'fontName' => 'Arial',
        'fontSize' => 11
    );

    // Información del documento word.
    $wordInfo   = Array(
        'creator' => 'Orlando Sergio Dominguez',
        'company' => 'Legislatura de Jujuy - Cómputos',
        'title' => 'Prueba de PHPWord',
        'description' => 'Documento de Ejemplo de una forma de implementación de la clase Office/PHPWord para crear documentos en formatos docx.',
        'category' => 'Implementación/clase/PHPWord',
        'modifiedBy' => '',
        'created' => Array(0, 0, 0, 3, 12, 2014),
        'modified' => Array(0, 0, 0, 3, 14, 2014),
        'subject' => 'PHPWord implementación',
        'keywords' => 'php,clase,PHPWord,implementación,ejemplo'
    );

    // Configuración creador de archivos word.
    $wordConfig = Array(
        "zoom"                  => 100,   // Int.
        "mirrorMargins"         => FALSE, // Bool.
        "decimalSymbol"         => ",",   //String (, | .).
        "themeFontLang"         => 'ES_ES', // String || NULL.

        // Corrector gramatical y ortografico.
        "hideSpellingErrors"    => TRUE,  // Bool.
        "hideGrammaticalErrors" => TRUE,  // Bool.
        "proofState"            => 'CLEAN', // String (DIRTY | CLEAN).

        // Seguimiento de cambios.
        "trackRevisions"        => TRUE,  // Bool.
        "doNotTrackMoves"       => TRUE,  // Bool.
        "doNotTrackFormatting"  => TRUE,  // Bool.

        //"revisionView"          => TRUE,  // Bool.
        // Proección del documento.
        "documentProtection"    => Array(TRUE,'READ_ONLY','password'),
        "updateFields"          => TRUE,  // Bool.
        
        // Separación de caractes.
        "autoHyphenation"       => TRUE,  // Bool.
        "consecutiveHyphenLimit"=> 0,     // Int.
        "hyphenationZone"       => 1,     // Int.
        "doNotHyphenateCaps"    => TRUE   // Bool.
     );

    $application = new Application(
        $pathHtml,
        $pathDocx
    );

    $application->createDocx(
        $default,
        $wordInfo,
        $wordConfig
    );

?>