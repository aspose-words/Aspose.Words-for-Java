/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */
package asposefeatures.workingwithdocument.insertpicture.java;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RelativeHorizontalPosition;
import com.aspose.words.RelativeVerticalPosition;
import com.aspose.words.WrapType;

public class AsposeInsertImage
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithdocument/insertpicture/data/";
		
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);

		builder.insertImage(dataPath + "background.jpg");
		builder.insertImage(dataPath + "background.jpg",
		        RelativeHorizontalPosition.MARGIN,
		        100,
		        RelativeVerticalPosition.MARGIN,
		        200,
		        200,
		        100,
		        WrapType.SQUARE);
		
		doc.save(dataPath + "AsposeImageInDoc.docx");
		
		System.out.println("Done.");
	}
}
