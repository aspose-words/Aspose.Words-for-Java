// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Hyphenation;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.Document;
import com.aspose.words.Run;
import com.aspose.words.WarningInfoCollection;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.words.WarningType;
import com.aspose.words.WarningSource;
import com.aspose.words.IHyphenationCallback;
import java.util.HashMap;
import com.aspose.ms.System.msConsole;


@Test
public class ExHyphenation extends ApiExampleBase
{
    @Test
    public void dictionary() throws Exception
    {
        //ExStart
        //ExFor:Hyphenation.IsDictionaryRegistered(String)
        //ExFor:Hyphenation.RegisterDictionary(String, String)
        //ExFor:Hyphenation.UnregisterDictionary(String)
        //ExSummary:Shows how to register a hyphenation dictionary.
        // A hyphenation dictionary contains a list of strings that define hyphenation rules for the dictionary's language.
        // When a document contains lines of text in which a word could be split up and continued on the next line,
        // hyphenation will look through the dictionary's list of strings for that word's substrings.
        // If the dictionary contains a substring, then hyphenation will split the word across two lines
        // by the substring and add a hyphen to the first half.
        // Register a dictionary file from the local file system to the "de-CH" locale.
        Hyphenation.registerDictionary("de-CH", getMyDir() + "hyph_de_CH.dic");

        Assert.assertTrue(Hyphenation.isDictionaryRegistered("de-CH"));

        // Open a document containing text with a locale matching that of our dictionary,
        // and save it to a fixed-page save format. The text in that document will be hyphenated.
        Document doc = new Document(getMyDir() + "German text.docx");

        Assert.That(doc.getFirstSection().getBody().getFirstParagraph().getRuns().<Run>OfType().All(
            r => r.Font.LocaleId == new CultureInfo("de-CH").LCID), assertTrue();

        doc.save(getArtifactsDir() + "Hyphenation.Dictionary.Registered.pdf");

        // Re-load the document after un-registering the dictionary,
        // and save it to another PDF, which will not have hyphenated text.
        Hyphenation.unregisterDictionary("de-CH");

        Assert.assertFalse(Hyphenation.isDictionaryRegistered("de-CH"));

        doc = new Document(getMyDir() + "German text.docx");
        doc.save(getArtifactsDir() + "Hyphenation.Dictionary.Unregistered.pdf");
        //ExEnd
    }

    @Test
    public void usePdfDocumentForDictionary() throws Exception
    {
        final String UNICODE_OPTIONAL_HYPHEN = "\u00ad";

        dictionary();

        Aspose.Pdf.Document pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Hyphenation.Dictionary.Registered.pdf");
        TextAbsorber textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);
        Assert.That(textAbsorber.Text.Replace("  ", " ").Contains($"La ob storen an deinen am sachen. Dop{unicodeOptionalHyphen}{Environment.NewLine}" +
            $"pelte um da am spateren verlogen ge{unicodeOptionalHyphen}{Environment.NewLine}" +
            $"kommen achtzehn blaulich."), assertTrue();

        pdfDoc = new Aspose.Pdf.Document(getArtifactsDir() + "Hyphenation.Dictionary.Unregistered.pdf");
        textAbsorber = new TextAbsorber();
        textAbsorber.Visit(pdfDoc);

        Assert.That(textAbsorber.Text.Replace("  ", " ").Contains($"La ob storen an deinen am sachen. {Environment.NewLine}" +
                                               $"Doppelte um da am spateren verlogen {Environment.NewLine}" +
                                               $"gekommen achtzehn blaulich."), assertTrue();
    }

    //ExStart
    //ExFor:Hyphenation
    //ExFor:Hyphenation.Callback
    //ExFor:Hyphenation.RegisterDictionary(String, Stream)
    //ExFor:Hyphenation.RegisterDictionary(String, String)
    //ExFor:Hyphenation.WarningCallback
    //ExFor:IHyphenationCallback
    //ExFor:IHyphenationCallback.RequestDictionary(String)
    //ExSummary:Shows how to open and register a dictionary from a file.
    @Test //ExSkip
    public void registerDictionary() throws Exception
    {
        // Set up a callback that tracks warnings that occur during hyphenation dictionary registration.
        WarningInfoCollection warningInfoCollection = new WarningInfoCollection();
        Hyphenation.setWarningCallback(warningInfoCollection);

        // Register an English (US) hyphenation dictionary by stream.
        Stream dictionaryStream = new FileStream(getMyDir() + "hyph_en_US.dic", FileMode.OPEN);
        Hyphenation.registerDictionaryInternal("en-US", dictionaryStream);

        Assert.assertEquals(0, warningInfoCollection.getCount());

        // Open a document with a locale that Microsoft Word may not hyphenate on an English machine, such as German.
        Document doc = new Document(getMyDir() + "German text.docx");

        // To hyphenate that document upon saving, we need a hyphenation dictionary for the "de-CH" language code.
        // This callback will handle the automatic request for that dictionary.
        Hyphenation.setCallback(new CustomHyphenationDictionaryRegister());

        // When we save the document, German hyphenation will take effect.
        doc.save(getArtifactsDir() + "Hyphenation.RegisterDictionary.pdf");

        // This dictionary contains two identical patterns, which will trigger a warning.
        Assert.assertEquals(1, warningInfoCollection.getCount());
        Assert.assertEquals(WarningType.MINOR_FORMATTING_LOSS, warningInfoCollection.get(0).getWarningType());
        Assert.assertEquals(WarningSource.LAYOUT, warningInfoCollection.get(0).getSource());
        Assert.assertEquals("Hyphenation dictionary contains duplicate patterns. The only first found pattern will be used. " +
                            "Content can be wrapped differently.", warningInfoCollection.get(0).getDescription());

        Hyphenation.setWarningCallback(null); //ExSkip
        Hyphenation.unregisterDictionary("en-US"); //ExSkip
        Hyphenation.setCallback(null); //ExSkip
    }

    /// <summary>
    /// Associates ISO language codes with local system filenames for hyphenation dictionary files.
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
                System.out.println(", is already registered.");
                return;
            }

            if (mHyphenationDictionaryFiles.containsKey(language))
            {
                Hyphenation.registerDictionary(language, mHyphenationDictionaryFiles.get(language));
                System.out.println(", successfully registered.");
                return;
            }

            System.out.println(", no respective dictionary file known by this Callback.");
        }

        private /*final*/ HashMap<String, String> mHyphenationDictionaryFiles;
    }
    //ExEnd
}

