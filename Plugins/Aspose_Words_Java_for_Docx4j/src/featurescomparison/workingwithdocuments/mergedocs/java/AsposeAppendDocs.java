/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */

package featurescomparison.workingwithdocuments.mergedocs.java;

import com.aspose.words.Document;
import com.aspose.words.ImportFormatMode;
import com.aspose.words.SaveFormat;

public class AsposeAppendDocs
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithdocuments/mergedocs/data/";
		
		Document doc1 = new Document(dataPath + "doc1.doc");
		Document doc2 = new Document(dataPath + "doc2.doc");
		
		doc1.appendDocument(doc2, ImportFormatMode.KEEP_SOURCE_FORMATTING);
		
		doc1.save(dataPath + "AsposeMerged.doc", SaveFormat.DOC);
	}
}
