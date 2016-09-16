package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.Style;
import com.aspose.words.StyleCollection;

public class AccessStyles {

	public static void main(String[] args) throws Exception {
		accessStyles();
		
		iterateThroughStyles();
	}
	
	public static void accessStyles() throws Exception {
		Document doc = new Document();
		StyleCollection styles = doc.getStyles();

		for (Style style : styles)
		    System.out.println(style.getName());
	}
	
	public static void iterateThroughStyles() throws Exception {
		Document doc = new Document();

		for(int i =0; i < doc.getStyles().getCount(); i++) 
		  System.out.println(doc.getStyles().get(i).getName());
	}
	
}
