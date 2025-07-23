package Examples;

// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class ExAI extends ApiExampleBase
{
    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiSummarize() throws Exception
    {
        //ExStart:AiSummarize
        //GistId:72d57eeddb7fb342fd51b26e5fcf9642
        //ExFor:GoogleAiModel
        //ExFor:OpenAiModel
        //ExFor:OpenAiModel.WithOrganization(String)
        //ExFor:OpenAiModel.WithProject(String)
        //ExFor:AiModel
        //ExFor:AiModel.Summarize(Document, SummarizeOptions)
        //ExFor:AiModel.Summarize(Document[], SummarizeOptions)
        //ExFor:AiModel.Create(AiModelType)
        //ExFor:AiModel.WithApiKey(String)
        //ExFor:AiModelType
        //ExFor:SummarizeOptions
        //ExFor:SummarizeOptions.#ctor
        //ExFor:SummarizeOptions.SummaryLength
        //ExFor:SummaryLength
        //ExSummary:Shows how to summarize text using OpenAI and Google models.
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
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:AiModel.Translate(Document, AI.Language)
        //ExFor:AI.Language
        //ExSummary:Shows how to translate text using Google models.
        Document doc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use Google generative language models.
        AiModel model = AiModel.create(AiModelType.GEMINI_15_FLASH).withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.ARABIC);
        translatedDoc.save(getArtifactsDir() + "AI.AiTranslate.docx");
        //ExEnd:AiTranslate
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiGrammar() throws Exception
    {
        //ExStart:AiGrammar
        //GistId:c012c14781944ce4cc5e31f35b08060a
        //ExFor:AiModel.CheckGrammar(Document, CheckGrammarOptions)
        //ExFor:CheckGrammarOptions
        //ExSummary:Shows how to check the grammar of a document.
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

    //ExStart:SelfHostedModel
    //GistId:67c1d01ce69d189983b497fd497a7768
    //ExFor:OpenAiModel
    //ExSummary:Shows how to use self-hosted AI model based on OpenAiModel.
    @Test (enabled = false, description = "This test should be run manually when you are configuring your model") //ExSkip
    public void selfHostedModel() throws Exception
    {
        Document doc = new Document(getMyDir() + "Big document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI generative language models.
        AiModel model = new CustomAiModel().withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.RUSSIAN);
        translatedDoc.save(getArtifactsDir() + "AI.SelfHostedModel.docx");
    }

    /// <summary>
    /// Custom self-hosted AI model.
    /// </summary>
    static class CustomAiModel extends OpenAiModel
    {
        /// <summary>
        /// Gets custom URL of the model.
        /// </summary>
        protected /*override*/ String getUrl() { return "https://localhost/"; }

        /// <summary>
        /// Gets model name.
        /// </summary>
        protected /*override*/ String getName() { return "my-model-24b"; }
    }
    //ExEnd:SelfHostedModel
}

