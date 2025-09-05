package DocsExamples.AI_powered_Features;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.Environment;
import com.aspose.words.AiModel;
import com.aspose.words.OpenAiModel;
import com.aspose.words.AiModelType;
import com.aspose.words.SummarizeOptions;
import com.aspose.words.SummaryLength;
import com.aspose.words.GoogleAiModel;
import com.aspose.words.Language;
import com.aspose.words.CheckGrammarOptions;


public class Working_with_AI extends DocsExamplesBase
{
    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiSummarize() throws Exception
    {
        //ExStart:AiSummarize
        //GistId:1e379bedb2b759c1be24c64aad54d13d
        Document firstDoc = new Document(getMyDir() + "Big document.docx");
        Document secondDoc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI or Google generative language models.
        AiModel model = ((OpenAiModel)AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey)).withOrganization("Organization").withProject("Project");

        SummarizeOptions options = new SummarizeOptions();

        options.setSummaryLength(SummaryLength.SHORT);
        Document oneDocumentSummary = model.summarize(firstDoc, options);
        oneDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.One.docx");

        options.setSummaryLength(SummaryLength.LONG);
        Document multiDocumentSummary = model.summarize(new Document[] { firstDoc, secondDoc }, options);
        multiDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.Multi.docx");
        //ExEnd:AiSummarize
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiTranslate() throws Exception
    {
        //ExStart:AiTranslate
        //GistId:ea14b3e44c0233eecd663f783a21c4f6
        Document doc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use Google generative language models.
        AiModel model = (GoogleAiModel)AiModel.create(AiModelType.GEMINI_15_FLASH).withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.ARABIC);
        translatedDoc.save(getArtifactsDir() + "AI.AiTranslate.docx");
        //ExEnd:AiTranslate
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiGrammar() throws Exception
    {
        //ExStart:AiGrammar
        //GistId:98a646d19cd7708ed0cd3d97b993a053
        Document doc = new Document(getMyDir() + "Big document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI generative language models.
        AiModel model = AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey);

        CheckGrammarOptions grammarOptions = new CheckGrammarOptions();
        grammarOptions.setImproveStylistics(true);

        Document proofedDoc = model.checkGrammar(doc, grammarOptions);
        proofedDoc.save(getArtifactsDir() + "AI.AiGrammar.docx");
        //ExEnd:AiGrammar
    }
}

