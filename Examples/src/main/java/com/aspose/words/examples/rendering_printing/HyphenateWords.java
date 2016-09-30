package com.aspose.words.examples.rendering_printing;

import java.io.FileInputStream;
import java.io.InputStream;

import com.aspose.words.Document;
import com.aspose.words.Hyphenation;
import com.aspose.words.examples.Utils;

public class HyphenateWords {

	private static final String dataDir = Utils.getSharedDataDir(HyphenateWords.class) + "RenderingAndPrinting/";

	public static void main(String[] args) throws Exception {
		//  Load hyphenation dictionaries for a specified languages from file.
		loadHyphenationDictionaryFromFile();
		
		// Load a hyphenation dictionary for a specified language from a stream.
		loadHyphenationDictionaryFromStream();
	}

	public static void loadHyphenationDictionaryFromFile() throws Exception {
		Document doc = new Document(dataDir + "in.docx");

		Hyphenation.registerDictionary("en-US", dataDir + "hyph_en_US.dic");
		Hyphenation.registerDictionary("de-CH", dataDir + "hyph_de_CH.dic");

		doc.save(dataDir + "LoadHyphenationDictionaryFromFile_Out.pdf");

	}

	public static void loadHyphenationDictionaryFromStream() throws Exception {
		Document doc = new Document(dataDir + "in.docx");

		InputStream stream = new FileInputStream(dataDir + "hyph_de_CH.dic");
		Hyphenation.registerDictionary("de-CH", stream);

		doc.save(dataDir + "LoadHyphenationDictionaryFromStream_Out.pdf");

	}

}
