package com.aspose.words.examples.loading_saving;

import com.aspose.words.Document;
import com.aspose.words.examples.Utils;

public class LoadTxt {
	public static void main(String[] args) throws Exception {

		// ExStart:LoadText
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(LoadTxt.class);

		// The encoding of the text file is automatically detected.
		Document doc = new Document(dataDir + "LoadTxt.txt");

		// Save as any Aspose.Words supported format, such as DOCX.
		doc.save(dataDir + "output.docx");
		// ExEnd:LoadText

		System.out.println("Loaded data from text file successfully.");
	}
}
