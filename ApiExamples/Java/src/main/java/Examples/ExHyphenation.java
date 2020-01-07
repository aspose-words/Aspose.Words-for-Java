package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.util.HashMap;

public class ExHyphenation extends ApiExampleBase {
    //ExStart
    //ExFor:Hyphenation
    //ExFor:Hyphenation.Callback
    //ExFor:Hyphenation.RegisterDictionary(String, Stream)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.WarningCallback
    //ExFor:IHyphenationCallback
    //ExFor:IHyphenationCallback.RequestDictionary(System.String)
    //ExSummary:Shows how to open and register a dictionary from a file.
    @Test //ExSkip
    public void registerDictionary() throws Exception {
        // Set up a callback that tracks warnings that occur during hyphenation dictionary registration
        WarningInfoCollection warningInfoCollection = new WarningInfoCollection();
        Hyphenation.setWarningCallback(warningInfoCollection);

        // Register an English (US) hyphenation dictionary by stream
        FileInputStream dictionaryStream = new FileInputStream(getMyDir() + "hyph_en_US.dic");
        Hyphenation.registerDictionary("en-US", dictionaryStream);

        // No warnings detected
        Assert.assertEquals(warningInfoCollection.getCount(), 0);

        // Open a document with a German locale that might not get automatically hyphenated by Microsoft Word an english machine
        Document doc = new Document(getMyDir() + "RandomGermanWords.doc");

        // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code
        // This callback will handle the automatic request for that dictionary
        Hyphenation.setCallback(new CustomHyphenationDictionaryRegister());

        // When we save the document, it will be hyphenated according to rules defined by the dictionary known by our callback
        doc.save(getArtifactsDir() + "Hyphenation.RegisterDictionary.pdf");

        // This dictionary contains two identical patterns, which will trigger a warning
        Assert.assertEquals(warningInfoCollection.getCount(), 1);
        Assert.assertEquals(warningInfoCollection.get(0).getWarningType(), WarningType.MINOR_FORMATTING_LOSS);
        Assert.assertEquals(warningInfoCollection.get(0).getSource(), WarningSource.LAYOUT);
        Assert.assertEquals(warningInfoCollection.get(0).getDescription(), "Hyphenation dictionary contains duplicate patterns. " +
                "The only first found pattern will be used. Content can be wrapped differently.");
    }

    /// <summary>
    /// Associates ISO language codes with custom local system dictionary files for their respective languages
    /// </summary>
    private static class CustomHyphenationDictionaryRegister implements IHyphenationCallback {
        public CustomHyphenationDictionaryRegister() {
            mHyphenationDictionaryFiles = new HashMap<>();
            {
                mHyphenationDictionaryFiles.put("en-US", getMyDir() + "hyph_en_US.dic");
                mHyphenationDictionaryFiles.put("de-CH", getMyDir() + "hyph_de_CH.dic");
            }
        }

        public void requestDictionary(String language) throws Exception {
            System.out.print("Hyphenation dictionary requested: " + language);

            if (Hyphenation.isDictionaryRegistered(language)) {
                System.out.println(", is already registered.");
                return;
            }

            if (mHyphenationDictionaryFiles.containsKey(language)) {
                Hyphenation.registerDictionary(language, mHyphenationDictionaryFiles.get(language));
                System.out.println(", successfully registered.");
                return;
            }

            System.out.println(", no respective dictionary file known by this Callback.");
        }

        private HashMap<String, String> mHyphenationDictionaryFiles;
    }
    //ExEnd

    @Test
    public void isDictionaryRegisteredEx() throws Exception {
        //ExStart
        //ExFor:Hyphenation.IsDictionaryRegistered(String)
        //ExSummary:Shows how to open check if some dictionary is registered.
        Document doc = new Document(getMyDir() + "Document.doc");
        Hyphenation.registerDictionary("en-US", getMyDir() + "hyph_en_US.dic");

        Assert.assertTrue(Hyphenation.isDictionaryRegistered("en-US"));
        //ExEnd
    }

    @Test
    public void unregisterDictionaryEx() throws Exception {
        //ExStart
        //ExFor:Hyphenation.UnregisterDictionary(String)
        //ExSummary:Shows how to un-register a dictionary
        Document doc = new Document(getMyDir() + "Document.doc");
        Hyphenation.registerDictionary("en-US", getMyDir() + "hyph_en_US.dic");

        Hyphenation.unregisterDictionary("en-US");

        Assert.assertFalse(Hyphenation.isDictionaryRegistered("en-US"));
        //ExEnd
    }
}
