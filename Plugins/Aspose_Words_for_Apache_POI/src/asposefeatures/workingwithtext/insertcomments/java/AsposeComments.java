/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */

package asposefeatures.workingwithtext.insertcomments.java;

import java.util.Date;

import com.aspose.words.Comment;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;

public class AsposeComments
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithtext/insertcomments/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.write("Some text is added.");

		Comment comment = new Comment(doc, "Aspose", "As", new Date());
		builder.getCurrentParagraph().appendChild(comment);
		comment.getParagraphs().add(new Paragraph(doc));
		comment.getFirstParagraph().getRuns().add(new Run(doc, "Comment text."));

		doc.save(dataPath + "AsposeComments.docx");
		
		System.out.println("Done.");
	}
}