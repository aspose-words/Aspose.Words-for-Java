package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class ReplaceInHeaderAndFooter {

    public static void main(String[] args) throws Exception {
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(ReplaceInHeaderAndFooter.class) + "FindAndReplace/";
        ReplaceTextInFooter(dataDir);

    }

    private static void ReplaceTextInFooter(String dataDir) throws Exception {
        // ExStart:ReplaceTextInFooter
        // Open the template document, containing obsolete copyright information in the footer.
        Document doc = new Document(dataDir + "HeaderFooter.ReplaceText.doc");

        HeaderFooterCollection headersFooters = doc.getFirstSection().getHeadersFooters();
        HeaderFooter footer = headersFooters.get(HeaderFooterType.FOOTER_PRIMARY);

        FindReplaceOptions options = new FindReplaceOptions();
        options.setMatchCase(false);
        options.setFindWholeWordsOnly(false);

        footer.getRange().replace("(C) 2006 Aspose Pty Ltd.", "Copyright (C) 2019 by Aspose Pty Ltd.", options);

        doc.save(dataDir + "HeaderFooter.ReplaceText.doc");
        // ExEnd:ReplaceTextInFooter
    }


}
