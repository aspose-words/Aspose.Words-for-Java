package com.aspose.words.examples.rendering_printing;

import com.aspose.words.Document;
import com.aspose.words.Hyphenation;
import com.aspose.words.IHyphenationCallback;
import com.aspose.words.examples.Utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class HyphenateWords {

    private static final String dataDir = Utils.getSharedDataDir(HyphenateWords.class) + "RenderingAndPrinting/";

    public static void main(String[] args) throws Exception {
        //  Load hyphenation dictionaries for a specified languages from file.
        loadHyphenationDictionaryFromFile();

        // Load a hyphenation dictionary for a specified language from a stream.
        loadHyphenationDictionaryFromStream();

        hyphenationCallback();
    }

    public static void loadHyphenationDictionaryFromFile() throws Exception {
        //ExStart:LoadHyphenationDictionaryFromFile
        Document doc = new Document(dataDir + "in.docx");

        Hyphenation.registerDictionary("en-US", dataDir + "hyph_en_US.dic");
        Hyphenation.registerDictionary("de-CH", dataDir + "hyph_de_CH.dic");

        doc.save(dataDir + "LoadHyphenationDictionaryFromFile_Out.pdf");
        //ExEnd:LoadHyphenationDictionaryFromFile
    }

    public static void loadHyphenationDictionaryFromStream() throws Exception {
        //ExStart:LoadHyphenationDictionaryFromStream
        Document doc = new Document(dataDir + "in.docx");

        InputStream stream = new FileInputStream(dataDir + "hyph_de_CH.dic");
        Hyphenation.registerDictionary("de-CH", stream);

        doc.save(dataDir + "LoadHyphenationDictionaryFromStream_Out.pdf");
        //ExEnd:LoadHyphenationDictionaryFromStream
    }

    //ExStart:HyphenationCallback
    public static void hyphenationCallback() {
        try {
            // Register hyphenation callback.
            Hyphenation.setCallback(new CustomHyphenationCallback());

            Document document = new Document(dataDir + "in.docx");
            document.save(dataDir + "LoadHyphenationDictionaryFromStream_Out.pdf");
        } catch (Exception e) {
            System.out.println(e.getMessage());
        } finally {
            Hyphenation.setCallback(null);
        }
    }

    static class CustomHyphenationCallback implements IHyphenationCallback {
        public void requestDictionary(String language) throws Exception {
            String dictionaryFolder = dataDir;
            String dictionaryFullFileName;
            switch (language) {
                case "en-US":
                    dictionaryFullFileName = new File(dictionaryFolder, "hyph_en_US.dic").getPath();
                    break;
                case "de-CH":
                    dictionaryFullFileName = new File(dictionaryFolder, "hyph_de_CH.dic").getPath();
                    break;
                default:
                    throw new Exception("Missing hyphenation dictionary for " + language);
            }
            // Register dictionary for requested language.
            Hyphenation.registerDictionary(language, dictionaryFullFileName);
        }
    }
    //ExEnd:HyphenationCallback
}
