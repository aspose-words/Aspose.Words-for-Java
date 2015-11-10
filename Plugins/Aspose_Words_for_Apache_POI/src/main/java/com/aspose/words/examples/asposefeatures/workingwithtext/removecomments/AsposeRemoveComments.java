package com.aspose.words.examples.asposefeatures.workingwithtext.removecomments;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.examples.Utils;

public class AsposeRemoveComments
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeRemoveComments.class);

	Document doc = new Document(dataDir + "AsposeComments.docx");
	
	// Collect all comments in the document
	NodeCollection comments = doc.getChildNodes(NodeType.COMMENT, true);
	// Look through all comments and remove those written by the authorName author.
	for (int i = comments.getCount() - 1; i >= 0; i--)
	{
	    Comment comment = (Comment) comments.get(i);
	    if (comment.getAuthor().equalsIgnoreCase("Aspose"))
		System.out.println("Aspose comment removed");
		comment.remove();
	}
	
	doc.save(dataDir + "AsposeCommentsRemoved.docx");
	System.out.println("Done...");
    }
}