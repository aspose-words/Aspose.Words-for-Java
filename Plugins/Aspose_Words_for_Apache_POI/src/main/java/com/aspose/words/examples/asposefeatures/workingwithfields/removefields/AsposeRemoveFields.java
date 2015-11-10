package com.aspose.words.examples.asposefeatures.workingwithfields.removefields;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Field;
import com.aspose.words.examples.Utils;

public class AsposeRemoveFields
{
    public static void main(String[] args) throws Exception
    {
	// The path to the documents directory.
        String dataDir = Utils.getDataDir(AsposeRemoveFields.class);

	Document doc = new Document();
	
	DocumentBuilder builder = new DocumentBuilder(doc);
	
	Field field = builder.insertField("PAGE");
	
	// Calling this method completely removes the field from the document.
	field.remove();
	
	doc.save(dataDir + "AsposeFieldsRemoved.docx");
	System.out.println("Aspose Fields Removed...");
    }
}
