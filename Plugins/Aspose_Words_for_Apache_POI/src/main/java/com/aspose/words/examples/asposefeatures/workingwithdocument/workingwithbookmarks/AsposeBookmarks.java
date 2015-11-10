package com.aspose.words.examples.asposefeatures.workingwithdocument.workingwithbookmarks;

import com.aspose.words.Bookmark;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeBookmarks
{
    // See more @ http://www.aspose.com/docs/display/wordsjava/Bookmarks+in+Aspose.Words
    
    public static void main(String[] args) throws Exception
    {
        // The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeBookmarks.class);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startBookmark("AsposeBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("AsposeBookmark");

        // By index.
        Bookmark bookmark1 = doc.getRange().getBookmarks().get(0);

        // By name.
        Bookmark bookmark2 = doc.getRange().getBookmarks().get("AsposeBookmark");

        doc.save(dataDir + "AsposeBookmark.doc", SaveFormat.DOC);

        System.out.println("Done.");
    }
}
