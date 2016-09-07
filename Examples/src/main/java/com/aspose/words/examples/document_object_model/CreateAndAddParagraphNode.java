package com.aspose.words.examples.document_object_model;

import com.aspose.words.Document;
import com.aspose.words.Paragraph;
import com.aspose.words.Section;

public class CreateAndAddParagraphNode {

	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		Paragraph para = new Paragraph(doc);
		Section section = doc.getLastSection();
		section.getBody().appendChild(para);
	}
}
