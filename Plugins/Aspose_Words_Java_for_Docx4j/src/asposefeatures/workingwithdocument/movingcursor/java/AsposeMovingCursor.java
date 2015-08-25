/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */

package asposefeatures.workingwithdocument.movingcursor.java;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Node;
import com.aspose.words.Paragraph;

public class AsposeMovingCursor
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithdocument/movingcursor/data/";
		
		Document doc = new Document(dataPath + "document.doc");
		DocumentBuilder builder = new DocumentBuilder(doc);

		//Shows how to access the current node in a document builder.
		Node curNode = builder.getCurrentNode();
		Paragraph curParagraph = builder.getCurrentParagraph();
		
		// Shows how to move a cursor position to a specified node.
		builder.moveTo(doc.getFirstSection().getBody().getLastParagraph());
		
		// Shows how to move a cursor position to the beginning or end of a document.
		builder.moveToDocumentEnd();
		builder.writeln("This is the end of the document.");

		builder.moveToDocumentStart();
		builder.writeln("This is the beginning of the document.");
		
		doc.save(dataPath + "AsposeMovingCursor.doc");
		
		System.out.println("Done.");
	}
}


