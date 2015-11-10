package com.aspose.words.examples.asposefeatures.workingwithtext.extractcomments;

import java.util.ArrayList;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.SaveFormat;
import com.aspose.words.examples.Utils;

public class AsposeExtractComments
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeExtractComments.class);

	Document doc = new Document(dataDir + "AsposeComments.docx");

	ArrayList collectedComments = new ArrayList();
	
	// Collect all comments in the document
	NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
	
	// Look through all comments and gather information about them.
	for (Comment comment : (Iterable<Comment>) comments)
	{
	    System.out.println(comment.getAuthor() + " - " + comment.getDateTime() + " - "
		    + comment.toString(SaveFormat.TEXT));
	}
	System.out.println("Done.");
    }
}
