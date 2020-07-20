package com.aspose.words.examples.programming_documents.find_replace;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

/**
 * Created by awaishafeez on 1/3/2018.
 */
public class ReplaceHtmlTextWithMeta_Characters {
    public static void main(String[] args) throws Exception {
        // ExStart:ReplaceHtmlTextWithMetaCharacters
        // The path to the documents directory.
        String dataDir = Utils.getSharedDataDir(ReplaceTextWithField.class) + "FindAndReplace/";
        String html = "<p>&ldquo;Some Text&rdquo;</p>";

        // Initialize a Document.
        Document doc = new Document();

        // Use a document builder to add content to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.write("{PLACEHOLDER}");

        FindReplaceOptions findReplaceOptions = new FindReplaceOptions();
        findReplaceOptions.setReplacingCallback(new FindAndInsertHtml());

        doc.getRange().replace("{PLACEHOLDER}", html, findReplaceOptions);

        dataDir = dataDir + "ReplaceHtmlTextWithMetaCharacters_out.doc";
        doc.save(dataDir);
        // ExEnd:ReplaceHtmlTextWithMetaCharacters
        System.out.println("\nText replaced with meta characters successfully.\nFile saved at " + dataDir);
    }

    // ExStart:ReplaceHtmlFindAndInsertHtml
    static class FindAndInsertHtml implements IReplacingCallback {
        public int replacing(ReplacingArgs e) throws Exception {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.getMatchNode();
            // create Document Buidler and insert MergeField
            DocumentBuilder builder = new DocumentBuilder((Document) e.getMatchNode().getDocument());
            builder.moveTo(currentNode);
            builder.insertHtml(e.getReplacement());
            currentNode.remove();
            //Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.SKIP;
        }
    }
    // ExEnd:ReplaceHtmlFindAndInsertHtml
}
