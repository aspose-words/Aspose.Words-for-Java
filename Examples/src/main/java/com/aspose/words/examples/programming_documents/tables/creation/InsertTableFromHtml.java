package com.aspose.words.examples.programming_documents.tables.creation;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class InsertTableFromHtml {

	private static final String dataDir = Utils.getSharedDataDir(InsertTableFromHtml.class) + "Tables/";

	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		
		// Insert the table from HTML. Note that AutoFitSettings does not apply to tables
		// inserted from HTML.
		builder.insertHtml("<table>" + 
				"<tr>" + 
					"<td>Row 1, Cell 1</td>" + 
					"<td>Row 1, Cell 2</td>" + 
				"</tr>" + 
				"<tr>" + 
					"<td>Row 2, Cell 2</td>" + 
					"<td>Row 2, Cell 2</td>" + 
				"</tr>" + 
				"</table>");

		// Save the document to disk.
		doc.save(dataDir + "DocumentBuilder_InsertTableFromHtml_Out.doc");
	}
}