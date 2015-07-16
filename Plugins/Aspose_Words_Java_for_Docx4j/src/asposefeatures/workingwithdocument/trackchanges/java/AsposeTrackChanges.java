/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */

package asposefeatures.workingwithdocument.trackchanges.java;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

public class AsposeTrackChanges
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithdocument/trackchanges/data/";
		
		Document doc = new Document(dataPath +"trackDoc.doc");
		doc.acceptAllRevisions();
		doc.save(dataPath + "AsposeAcceptChanges.doc", SaveFormat.DOC);
		
		System.out.println("Done.");
	}
}
