<?php

/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use java\io\File as File;
use java\text\MessageFormat as MessageFormat ;
use com\aspose\words\Document as Document;
use com\aspose\words\LayoutCollector as LayoutCollector;

class PageSplitter {

    public static function main() {

        // The path to the documents directory.
        $dataDir = "/usr/local/apache-tomcat-8.0.22/webapps/JavaBridge/Aspose_Words_Java_For_PHP/src/loadingandsaving/pagesplitter/documentpagesplitter/data/";

        PageSplitter::SplitAllDocumentsToPages($dataDir);

    }

    public static function SplitAllDocumentsToPages($folderName) {

        $files_obj = new File($folderName);
        $files = $files_obj->listFiles();

        foreach ($files as $file) {
            if (java_values($file->isFile())) {
                PageSplitter::SplitDocumentToPages($file);
            }
        }

    }

    public static function SplitDocumentToPages(File $docName) {

        $folderName = $docName->getParent();
        $fileName = $docName->getName();
        $extensionName = $fileName->substring($fileName->lastIndexOf("."));
        $outFolder_obj = new File($folderName, "Out");
        $outFolder = $outFolder_obj->getAbsolutePath();

        echo "<BR> Processing document: " . $fileName . $extensionName;

        $doc = new Document($docName->getAbsolutePath());

        // Create and attach collector to the document before page layout is built.
        $layoutCollector = new LayoutCollector(doc);

        // This will build layout model and collect necessary information.
        $doc->updatePageLayout();

        // Split nodes in the document into separate pages.
        $splitter = new DocumentPageSplitter($layoutCollector);

        // Save each page to the disk as a separate document.
        for ($page = 1; $page <= java_values($doc->getPageCount()); $page++)
        {
            $pageDoc = $splitter->GetDocumentOfPage($page);
            $file_obj = new File($outFolder, MessageFormat::format("{0} - page{1} Out{2}", $fileName, $page, $extensionName));
            $abs_path = $file_obj->getAbsolutePath();
            $pageDoc->save($abs_path);
        }

        // Detach the collector from the document.
        $layoutCollector->setDocument(null);

    }

}