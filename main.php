<?php

    require 'vendor/autoload.php';
    require 'src/class.application.base.php';
    require 'src/class.application.php';

    $pathHtml = $argv[1];
    $pathDocx = $argv[2];

    $application = new Application($pathHtml,$pathDocx);
    $application->createDocx();

?>