package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.java.collections.StringSwitchMap;
import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.Hyphenation;
import com.aspose.ms.System.IO.Stream;
import java.io.FileInputStream;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.msConsole;
import com.aspose.words.IHyphenationCallback;
import com.aspose.ms.System.IO.Path;


class WorkingWithHyphenation extends DocsExamplesBase
{
    @Test
    public void hyphenateWordsOfLanguages() throws Exception
    {
        //ExStart:HyphenateWordsOfLanguages
        Document doc = new Document(getMyDir() + "German text.docx");

        Hyphenation.registerDictionary("en-US", getMyDir() + "hyph_en_US.dic");
        Hyphenation.registerDictionary("de-CH", getMyDir() + "hyph_de_CH.dic");

        doc.save(getArtifactsDir() + "WorkingWithHyphenation.HyphenateWordsOfLanguages.pdf");
        //ExEnd:HyphenateWordsOfLanguages
    }

    @Test
    public void loadHyphenationDictionaryForLanguage() throws Exception
    {
        //ExStart:LoadHyphenationDictionaryForLanguage
        Document doc = new Document(getMyDir() + "German text.docx");
        
        Stream stream = new FileInputStream(getMyDir() + "hyph_de_CH.dic");
        Hyphenation.registerDictionaryInternal("de-CH", stream);

        doc.save(getArtifactsDir() + "WorkingWithHyphenation.LoadHyphenationDictionaryForLanguage.pdf");
        //ExEnd:LoadHyphenationDictionaryForLanguage
    }

    @Test 
    //ExStart:CustomHyphenation
    public void hyphenationCallback() throws Exception
    {
        try
        {
            // Register hyphenation callback.
            Hyphenation.setCallback(new CustomHyphenationCallback());

            Document document = new Document(getMyDir() + "German text.docx");
            document.save(getArtifactsDir() + "WorkingWithHyphenation.HyphenationCallback.pdf");
        }
        catch (Exception e) when (e.Message.StartsWith("Missing hyphenation dictionary"))
        {
            msConsole.WriteLine(e.Message);
        }
        finally
        {
            Hyphenation.setCallback(null);
        }
    }

    public static class CustomHyphenationCallback implements IHyphenationCallback
    {
        public void requestDictionary(String language) throws Exception
        {
            String dictionaryFolder = getMyDir();
            String dictionaryFullFileName;
            switch (gStringSwitchMap.of(language))
            {
                case /*"en-US"*/0:
                    dictionaryFullFileName = Path.combine(dictionaryFolder, "hyph_en_US.dic");
                    break;
                case /*"de-CH"*/1:
                    dictionaryFullFileName = Path.combine(dictionaryFolder, "hyph_de_CH.dic");
                    break;
                default:
                    throw new Exception($"Missing hyphenation dictionary for {language}.");
            }
            // Register dictionary for requested language.
            Hyphenation.registerDictionary(language, dictionaryFullFileName);
        }
    }

	//JAVA-added for string switch emulation
	private static final StringSwitchMap gStringSwitchMap = new StringSwitchMap
	(
		"en-US",
		"de-CH"
	);

    //ExEnd:CustomHyphenation
}
