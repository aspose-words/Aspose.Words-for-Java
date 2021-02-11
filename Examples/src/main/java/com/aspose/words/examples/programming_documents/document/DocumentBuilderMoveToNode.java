package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.NodeType;
import org.testng.Assert;

public class DocumentBuilderMoveToNode {
    public static void main(String[] args) throws Exception {
        //ExStart:DocumentBuilderMoveToNode
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a bookmark and add content to it using a DocumentBuilder.
        builder.startBookmark("MyBookmark");
        builder.writeln("Bookmark contents.");
        builder.endBookmark("MyBookmark");

        // The node that the DocumentBuilder is currently at is past the boundaries of the bookmark.
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkEnd(), builder.getCurrentParagraph().getFirstChild());

        // If we wish to revise the content of our bookmark with the DocumentBuilder, we can move back to it like this.
        builder.moveToBookmark("MyBookmark");

        // Now we're located between the bookmark's BookmarkStart and BookmarkEnd nodes, so any text the builder adds will be within it.
        Assert.assertEquals(doc.getRange().getBookmarks().get(0).getBookmarkStart(), builder.getCurrentParagraph().getFirstChild());

        // We can move the builder to an individual node,
        // which in this case will be the first node of the first paragraph, like this.
        builder.moveTo(doc.getFirstSection().getBody().getFirstParagraph().getChildNodes(NodeType.ANY, false).get(0));

        Assert.assertEquals(NodeType.BOOKMARK_START, builder.getCurrentNode().getNodeType());
        Assert.assertTrue(builder.isAtStartOfParagraph());

        // A shorter way of moving the very start/end of a document is with these methods.
        builder.moveToDocumentEnd();
        Assert.assertTrue(builder.isAtEndOfParagraph());
        builder.moveToDocumentStart();
        Assert.assertTrue(builder.isAtStartOfParagraph());
        //ExEnd:DocumentBuilderMoveToNode
    }
}