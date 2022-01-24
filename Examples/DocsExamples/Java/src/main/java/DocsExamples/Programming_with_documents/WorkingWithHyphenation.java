package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.Document;
import com.aspose.words.Hyphenation;
import com.aspose.words.IHyphenationCallback;
import org.testng.annotations.Test;

import java.io.FileInputStream;
import java.nio.file.Paths;
import java.text.MessageFormat;

@Test
public class WorkingWithHyphenation extends DocsExamplesBase
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
        
        FileInputStream stream = new FileInputStream(getMyDir() + "hyph_de_CH.dic");
        Hyphenation.registerDictionary("de-CH", stream);

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
        catch (Exception e)
        {
            if (e.getMessage().startsWith("Missing hyphenation dictionary")) {
            System.out.println(e.getMessage());
        }

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
            switch (language)
            {
                case "en-US":
                    dictionaryFullFileName = Paths.get(dictionaryFolder, "hyph_en_US.dic").toString();
                    break;
                case "de-CH":
                    dictionaryFullFileName = Paths.get(dictionaryFolder, "hyph_de_CH.dic").toString();
                    break;
                default:
                    throw new Exception(MessageFormat.format("Missing hyphenation dictionary for {0}.", language));
            }
            // Register dictionary for requested language.
            Hyphenation.registerDictionary(language, dictionaryFullFileName);
        }
    }
    //ExEnd:CustomHyphenation
}
