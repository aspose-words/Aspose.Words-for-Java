<?php
/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

class CheckFormat {

    public static function check()
    {
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/loadingandsaving/checkformat/data/";
        $supportedDir = $dataDir . '/OutSupported/';
        $fileObj = new Java("java.io.File",$dataDir);
        $filesList  = $fileObj->listFiles();

        $loadFormat = java('com.aspose.words.LoadFormat');



        foreach($filesList as $file)
        {

            if(java_values($file->isDirectory()))
            {
                continue;
            }

            $nameOnly  = $file->getName();
            echo $nameOnly . '<br/>';

            $fileName = $file->getPath();
            echo $fileName . '<br/>';
            $infoObj = new Java('com.aspose.words.FileFormatUtil');
            $info = $infoObj->detectFileFormat($fileName);


            switch(java_values($info->getLoadFormat()))
            {
                case java_values($loadFormat->DOC):
                    echo ("Microsoft Word 97-2003 document.");
                    break;
                case java_values($loadFormat->DOT):
                    echo ("Microsoft Word 97-2003 template.");
                    break;
                case java_values($loadFormat->DOCX):
                    echo ("Office Open XML WordprocessingML Macro-Free Document.");
                    break;
                case java_values($loadFormat->DOCM):
                    echo ("Office Open XML WordprocessingML Macro-Enabled Document.");
                    break;
                case java_values($loadFormat->DOTX):
                    echo ("Office Open XML WordprocessingML Macro-Free Template.");
                    break;
                case java_values($loadFormat->DOTM):
                    echo ("Office Open XML WordprocessingML Macro-Enabled Template.");
                    break;
                case java_values($loadFormat->FLAT_OPC):
                    echo ("Flat OPC document.");
                    break;
                case java_values($loadFormat->RTF):
                    echo ("RTF format.");
                    break;
                case java_values($loadFormat->WORD_ML):
                    echo ("Microsoft Word 2003 WordprocessingML format.");
                    break;
                case java_values($loadFormat->HTML):
                    echo ("HTML format.");
                    break;
                case java_values($loadFormat->MHTML):
                    echo ("MHTML (Web archive) format.");
                    break;
                case java_values($loadFormat->ODT):
                    echo ("OpenDocument Text.");
                    break;
                case java_values($loadFormat->OTT):
                    echo ("OpenDocument Text Template.");
                    break;
                case java_values($loadFormat->DOC_PRE_WORD_97):
                    echo ("MS Word 6 or Word 95 format.");
                    break;
                case java_values($loadFormat->UNKNOWN):
                default:
                    echo ("Unknown format.");
                    break;
            }
            echo '<br/>';
            $destFileObj = new Java("java.io.File",$supportedDir . $nameOnly);
            $destFile = $destFileObj->getPath();
            copy(java_values($fileName), java_values($destFile));
        }
    }
}

CheckFormat::check();
