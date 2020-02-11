// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.WarningInfoCollection;
import com.aspose.words.Hyphenation;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.NUnit.Framework.msAssert;
import org.testng.Assert;
import com.aspose.words.Document;
import com.aspose.words.WarningType;
import com.aspose.words.WarningSource;
import com.aspose.words.IHyphenationCallback;
import java.util.HashMap;
import com.aspose.ms.System.msConsole;


@Test
public class ExHyphenation extends ApiExampleBase
{
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
    public void registerDictionary() throws Exception
    {
        // Set up a callback that tracks warnings that occur during hyphenation dictionary registration
        WarningInfoCollection warningInfoCollection = new WarningInfoCollection();
        Hyphenation.setWarningCallback(warningInfoCollection);

        // Register an English (US) hyphenation dictionary by stream
        Stream dictionaryStream = new FileStream(getMyDir() + "hyph_en_US.dic", FileMode.OPEN);
        Hyphenation.registerDictionaryInternal("en-US", dictionaryStream);

        // No warnings detected
        msAssert.areEqual(0, warningInfoCollection.getCount());

        // Open a document with a German locale that might not get automatically hyphenated by Microsoft Word an english machine
        Document doc = new Document(getMyDir() + "Unhyphenated German text.docx");

        // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code
        // This callback will handle the automatic request for that dictionary 
        Hyphenation.setCallback(new CustomHyphenationDictionaryRegister());

        // When we save the document, it will be hyphenated according to rules defined by the dictionary known by our callback
        doc.save(getArtifactsDir() + "Hyphenation.RegisterDictionary.pdf");

        // This dictionary contains two identical patterns, which will trigger a warning
        msAssert.areEqual(1, warningInfoCollection.getCount());
        msAssert.areEqual(WarningType.MINOR_FORMATTING_LOSS, warningInfoCollection.get(0).getWarningType());
        msAssert.areEqual(WarningSource.LAYOUT, warningInfoCollection.get(0).getSource());
        msAssert.areEqual("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
                        "Content can be wrapped differently.", warningInfoCollection.get(0).getDescription());
    }

    /// <summary>
    /// Associates ISO language codes with custom local system dictionary files for their respective languages
    /// </summary>
    private static class CustomHyphenationDictionaryRegister implements IHyphenationCallback
    {
        public CustomHyphenationDictionaryRegister()
        {
            mHyphenationDictionaryFiles = new HashMap<String, String>();
            {
                mHyphenationDictionaryFiles.put( "en-US", getMyDir() + "hyph_en_US.dic");
                mHyphenationDictionaryFiles.put( "de-CH", getMyDir() + "hyph_de_CH.dic");
            }
        }

        public void requestDictionary(String language) throws Exception
        {
            msConsole.write("Hyphenation dictionary requested: " + language);

            if (Hyphenation.isDictionaryRegistered(language))
            {
                msConsole.writeLine(", is already registered.");
                return;
            }

            if (mHyphenationDictionaryFiles.containsKey(language))
            {
                Hyphenation.registerDictionary(language, mHyphenationDictionaryFiles.get(language));
                msConsole.writeLine(", successfully registered.");
                return;
            }

            msConsole.writeLine(", no respective dictionary file known by this Callback.");
        }

        private /*final*/ HashMap<String, String> mHyphenationDictionaryFiles;
    }
    //ExEnd

    @Test
    public void isDictionaryRegistered() throws Exception
    {
        //ExStart
        //ExFor:Hyphenation.IsDictionaryRegistered(String)
        //ExSummary:Shows how to open check if some dictionary is registered.
        Document doc = new Document(getMyDir() + "Document.docx");
        Hyphenation.registerDictionary("en-US", getMyDir() + "hyph_en_US.dic");

        msAssert.areEqual(true, Hyphenation.isDictionaryRegistered("en-US"));
        //ExEnd
    }

    @Test
    public void unregisteredDictionary() throws Exception
    {
        //ExStart
        //ExFor:Hyphenation.UnregisterDictionary(String)
        //ExSummary:Shows how to un-register a dictionary.
        Document doc = new Document(getMyDir() + "Document.docx");
        Hyphenation.registerDictionary("en-US", getMyDir() + "hyph_en_US.dic");

        Hyphenation.unregisterDictionary("en-US");

        msAssert.areEqual(false, Hyphenation.isDictionaryRegistered("en-US"));
        //ExEnd
    }
}
