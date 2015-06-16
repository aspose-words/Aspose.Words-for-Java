<?php

/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

use com\aspose\words\Document as Document;
use com\aspose\words\LayoutCollector as LayoutCollector;
use com\aspose\words\Node as Node;
use com\aspose\words\NodeType as NodeType;

use java\util\ArrayList as ArrayList;
use java\util\Hashtable as Hashtable;

class PageNumberFinder {

    // Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
    private $mNodeStartPageLookup;
    private $mNodeEndPageLookup;
        // Maps page number to a list of nodes found on that page.
    private $mReversePageLookup;
    private $mCollector;

    /**
     * Initializes new instance of this class.
     */

    public function __construct(LayoutCollector $collector) {

        $this->mNodeStartPageLookup = new Hashtable();
        $this->$mNodeEndPageLookup = new Hashtable();

        $this->mCollector = $collector;

    }

    /**
    * Retrieves 1-based index of a page that the node begins on.
    */

    public function GetPage(Node $node){

        if (java_values($this->mNodeStartPageLookup->containsKey($node)))
            return $this->mNodeStartPageLookup->get($node);

        return $this->mCollector->getStartPageIndex($node);

    }

    /**
     * Retrieves 1-based index of a page that the node ends on.
     */

    public function GetPageEnd(Node $node) {

        if ($this->mNodeEndPageLookup->containsKey($node))
            return $this->mNodeEndPageLookup->get($node);

        return $this->mCollector->getEndPageIndex($node);

    }

    /**
     * Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
     */

    public function PageSpan(Node $node) {

        return $this->GetPageEnd($node) - $this->GetPage($node) + 1;

    }

    /**
     * Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
     */

    public function RetrieveAllNodesOnPages($startPage, $endPage, $nodeType) {

        if (java_values($startPage) < 1 || java_values($startPage) > javav_values($this->getDocument()->getPageCount()))
            throw new Exception("startPage");

        if (java_values($endPage) < 1 || java_values($endPage) > java_values($this->getDocument()->getPageCount()) || java_values($endPage) < java_values($startPage))
            throw new Exception("endPage");

        $this->CheckPageListsPopulated();

        $pageNodes = new ArrayList();

        for ($page = $startPage; $page <= $endPage; $page++) {
            // Some pages can be empty.
            if (!$this->mReversePageLookup->containsKey($page))
                continue;

            $nodes = $this->mReversePageLookup->get($page);

            foreach ($nodes as $node) {
                if (java_values($node->getParentNode()) != null && (($nodeType == NodeType::ANY) || ($nodeType == java_values($node->getNodeType()))) && !java_values($pageNodes->contains($node)))
                    $pageNodes->add($node);
            }
        }

        return $pageNodes;

    }
}