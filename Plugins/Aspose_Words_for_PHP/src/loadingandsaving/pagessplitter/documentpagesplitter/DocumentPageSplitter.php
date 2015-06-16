<?php

/*
 * Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\Document as Document;
use com\aspose\words\LayoutCollector as LayoutCollector;
use com\aspose\words\NodeType as NodeType;

class DocumentPageSplitter {

    private $mPageNumberFinder;

    /**
     * Initializes new instance of this class. This method splits the document into sections so that each page
     * begins and ends at a section boundary. It is recommended not to modify the document afterwards.
     */

    public function __construct(LayoutCollector $collector) {

        $this->mPageNumberFinder = new PageNumberFinder($collector);
        $this->mPageNumberFinder->SplitNodesAcrossPages();

    }

    /**
     * Gets the document of a page.
     */

    public function GetDocumentOfPage($pageIndex) {

        return $this->GetDocumentOfPageRange($pageIndex, $pageIndex);

    }

    /**
     * Gets the document of a page range.
     */

    public function GetDocumentOfPageRange($startIndex, $endIndex) {

        $result = $this->getDocument()->deepClone(false);

        $this->mPageNumberFinder->RetrieveAllNodesOnPages($startIndex, $endIndex, NodeType.SECTION);

        $sections = $this->mPageNumberFinder->toArray();

        foreach($sections as $section) {

            $result->appendChild($result->importNode($section, true));
        }



        return $result;

    }

    /**
     * Gets the document this instance works with.
     */

    public function getDocument() {

        return $this->mPageNumberFinder->getDocument();
    }


}