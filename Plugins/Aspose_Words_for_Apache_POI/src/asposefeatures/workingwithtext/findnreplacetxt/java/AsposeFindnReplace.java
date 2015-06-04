/**
 * Copyright (c) Aspose 2002-2014. All Rights Reserved.
 */

package asposefeatures.workingwithtext.findnreplacetxt.java;

import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;

// For more info please visit http://www.aspose.com/docs/display/wordsjava/Find+and+Replace+Overview
public class AsposeFindnReplace
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/asposefeatures/workingwithtext/findnreplacetxt/data/";
		
		Document doc = new Document(dataPath + "replaceDoc.doc");
		
		// Replaces all 'sad' and 'mad' occurrences with 'bad'
		doc.getRange().replace("sad", "bad", false, true); 
		
		// Replaces all 'sad' and 'mad' occurrences with 'bad'
		doc.getRange().replace(Pattern.compile("[s|m]ad"), "bad");
		
		doc.save(dataPath + "AsposeReplaced.doc", SaveFormat.DOC);
		
		System.out.println("Done.");
	}
}