/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package programmingwithdocuments.workingwithcomments.processcomments.java;

import java.util.ArrayList;
import java.io.File;
import java.net.URI;

import com.aspose.words.*;

@SuppressWarnings("unchecked")
public class ProcessComments
{
    public static void main(String[] args) throws Exception
    {
        // A sample infrastructure.
        String dataDir = "src/programmingwithdocuments/workingwithcomments/processcomments/data/";

        // Open the document.
        Document doc = new Document(dataDir + "TestFile.doc");

        //ExStart
        //ExId:ProcessComments_Main
        //ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
        // Extract the information about the comments of all the authors.
        for (String comment : (Iterable<String>) extractComments(doc))
            System.out.print(comment);

        // Remove comments by the "pm" author.
        removeComments(doc, "pm");
        System.out.println("Comments from \"pm\" are removed!");

        // Extract the information about the comments of the "ks" author.
        for (String comment : (Iterable<String>) extractComments(doc, "ks"))
            System.out.print(comment);

        // Remove all comments.
        removeComments(doc);
        System.out.println("All comments are removed!");

        // Save the document.
        doc.save(dataDir + "Test File Out.doc");
        //ExEnd
    }

    //ExStart
    //ExFor:Comment.Author
    //ExFor:Comment.DateTime
    //ExId:ProcessComments_Extract_All
    //ExSummary:Extracts the author name, date&time and text of all comments in the document.
    static ArrayList extractComments(Document doc) throws Exception
    {
        ArrayList collectedComments = new ArrayList();
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and gather information about them.
        for (Comment comment : (Iterable<Comment>) comments)
        {
            collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " + comment.toString(SaveFormat.TEXT));
        }
        return collectedComments;
    }
    //ExEnd

    //ExStart
    //ExId:ProcessComments_Extract_Author
    //ExSummary:Extracts the author name, date&time and text of the comments by the specified author.
    static ArrayList extractComments(Document doc, String authorName) throws Exception
    {
        ArrayList collectedComments = new ArrayList();
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and gather information about those written by the authorName author.
        for (Comment comment : (Iterable<Comment>) comments)
        {
            if (comment.getAuthor().equals(authorName))
                collectedComments.add(comment.getAuthor() + " " + comment.getDateTime() + " " + comment.toString(SaveFormat.TEXT));
        }
        return collectedComments;
    }
    //ExEnd

    //ExStart
    //ExId:ProcessComments_Remove_All
    //ExSummary:Removes all comments in the document.
    static void removeComments(Document doc) throws Exception
    {
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Remove all comments.
        comments.clear();
    }
    //ExEnd

    //ExStart
    //ExId:ProcessComments_Remove_Author
    //ExSummary:Removes comments by the specified author.
    static void removeComments(Document doc, String authorName) throws Exception
    {
        // Collect all comments in the document
        NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
        // Look through all comments and remove those written by the authorName author.
        for (int i = comments.getCount() - 1; i >= 0; i--)
        {
            Comment comment = (Comment)comments.get(i);
            if (comment.getAuthor().equals(authorName))
                comment.remove();
        }
    }
    //ExEnd
}