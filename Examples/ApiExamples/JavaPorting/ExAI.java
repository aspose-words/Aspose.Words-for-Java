// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.Environment;
import com.aspose.words.AiModel;
import com.aspose.words.OpenAiModel;
import com.aspose.words.AiModelType;
import com.aspose.words.SummarizeOptions;
import com.aspose.words.SummaryLength;
import com.aspose.words.Language;
import com.aspose.words.CheckGrammarOptions;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.GoogleAiModel;
import com.aspose.words.AsposeLlmModel;


@Test
public class ExAI extends ApiExampleBase
{
    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiSummarize() throws Exception
    {
        //ExStart:AiSummarize
        //GistId:366eb64fd56dec3c2eaa40410e594182
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
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:AiModel.Translate(Document, AI.Language)
        //ExFor:AI.Language
        //ExSummary:Shows how to translate text using Google models.
        Document doc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use Google generative language models.
        AiModel model = AiModel.create(AiModelType.GEMINI_FLASH_LATEST).withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.ARABIC);
        translatedDoc.save(getArtifactsDir() + "AI.AiTranslate.docx");
        //ExEnd:AiTranslate
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiGrammar() throws Exception
    {
        //ExStart:AiGrammar
        //GistId:f86d49dc0e6781b93e576539a01e6ca2
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
    //ExFor:AiModel.Url
    //ExSummary:Shows how to use self-hosted AI model based on OpenAiModel.
    @Test (enabled = false, description = "This test should be run manually when you are configuring your model") //ExSkip
    public void selfHostedModel() throws Exception
    {
        Document doc = new Document(getMyDir() + "Big document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI generative language models.
        AiModel model = new CustomAiModel("my-model-24b", "https://my.a.com/").withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.RUSSIAN);
        translatedDoc.save(getArtifactsDir() + "AI.SelfHostedModel.docx");
    }

    /// <summary>
    /// Custom self-hosted AI model.
    /// </summary>
    static class CustomAiModel extends OpenAiModel
    {
        CustomAiModel(String name, String url)
        {
        	super(name);
	
            mUrl = url;
        }

        public /*override*/ String getUrl() { return mUrl; }

        private /*final*/ String mUrl;
    }
    //ExEnd:SelfHostedModel

    @Test
    public void changeDefaultUrl()
    {
        //ExStart:ChangeDefaultUrl
        //GistId:bd7947d9ad5eb092f532604cb15f593b
        //ExFor:AiModel.Url
        //ExSummary:Shows how to change model default url.
        String apiKey = System.getenv("API_KEY");
        AiModel model = AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey);
        // Default value "https://api.openai.com/".
        model.setUrl("https://my.a.com/");
        //ExEnd:ChangeDefaultUrl

        Assert.assertEquals("https://my.a.com/", model.getUrl());
    }

    @Test
    public void changeDefaultTimeout()
    {
        //ExStart:ChangeDefaultTimeout
        //GistId:bd7947d9ad5eb092f532604cb15f593b
        //ExFor:AiModel.Timeout
        //ExSummary:Shows how to change model default timeout.
        String apiKey = System.getenv("API_KEY");
        AiModel model = AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey);
        // Default value 100000ms.
        model.setTimeout(250000);
        //ExEnd:ChangeDefaultTimeout

        Assert.assertEquals(250000, model.getTimeout());
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void gemini() throws Exception
    {
        //ExStart:Gemini
        //GistId:0da8468118377c4860b28603bc95ffe6
        //ExFor:GoogleAiModel
        //ExFor:GoogleAiModel.#ctor(String)
        //ExFor:GoogleAiModel.#ctor(String, String)
        //ExSummary:Shows how to use google AI model.
        String apiKey = System.getenv("API_KEY");
        GoogleAiModel model = new GoogleAiModel("gemini-flash-latest", apiKey);

        Document doc = new Document(getMyDir() + "Big document.docx");
        SummarizeOptions summarizeOptions = new SummarizeOptions(); { summarizeOptions.setSummaryLength(SummaryLength.VERY_SHORT); }
        Document summary = model.summarize(doc, summarizeOptions);
        //ExEnd:Gemini
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void openAiModelConstructor() throws Exception
    {
        //ExStart:OpenAiModelConstructor
        //GistId:8c640b84550c83678329a9a92f10bcdd
        //ExFor:OpenAiModel.#ctor(String,String)
        //ExSummary:Shows how to create an OpenAI model instance directly using an API key and model name.
        String apiKey = System.getenv("API_KEY");
        // Create an OpenAI model instance using the constructor with model name and API key.
        OpenAiModel model = new OpenAiModel("gpt-4o-mini", apiKey);

        Document doc = new Document(getMyDir() + "Big document.docx");
        // Summarize the document using the OpenAI model with short summary length.
        SummarizeOptions summarizeOptions = new SummarizeOptions(); { summarizeOptions.setSummaryLength(SummaryLength.VERY_SHORT); }
        Document summary = model.summarize(doc, summarizeOptions);

        summary.save(getArtifactsDir() + "OpenAiModel.OpenAiModelConstructor.docx");
        //ExEnd:OpenAiModelConstructor

        // Verify the summary was generated (non-empty content).
        Assert.less(0, summary.getText().trim().length());
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void asposeLlmSummarize() throws Exception
    {
        //ExStart:AsposeLlmSummarize
        //GistId:13ac22ab31bbd694d29ab4db2a496206
        //ExFor:AsposeLlmModel.#ctor(String)
        //ExFor:AsposeLlmModel.Summarize(Document, SummarizeOptions)
        //ExFor:AsposeLlmModel.Summarize(Document[], SummarizeOptions)
        //ExFor:AsposeLlmModel.Dispose
        //ExSummary:Shows how to summarize documents with a local model, so the content never leaves the machine.
        // Aspose.LLM is a separate package with its own license, for example:
        // new Aspose.LLM.License().SetLicense("Aspose.Total.lic");
        Document firstDoc = new Document(getMyDir() + "Big document.docx");
        Document secondDoc = new Document(getMyDir() + "Document.docx");

        // Disposing the model releases the local model and its native resources.
        AsposeLlmModel model = new AsposeLlmModel("Qwen25_3BPresetCpu");
        try /*JAVA: was using*/
        {
            SummarizeOptions summarizeOptions = new SummarizeOptions();
            
            summarizeOptions.setSummaryLength(SummaryLength.SHORT);
            Document summary = model.summarize(firstDoc, summarizeOptions);
            summary.save(getArtifactsDir() + "AI.AsposeLlmSummarize.One.docx");

            summarizeOptions.setSummaryLength(SummaryLength.MEDIUM);
            Document combinedSummary = model.summarize(new Document[] { firstDoc, secondDoc }, summarizeOptions);
            combinedSummary.save(getArtifactsDir() + "AI.AsposeLlmSummarize.Multiple.docx");
        }
        finally { if (model != null) model.dispose(); }
        //ExEnd:AsposeLlmSummarize
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void asposeLlmTranslate() throws Exception
    {
        //ExStart:AsposeLlmTranslate
        //GistId:13ac22ab31bbd694d29ab4db2a496206
        //ExFor:AsposeLlmModel.Translate(Document, AI.Language)
        //ExSummary:Shows how to translate a document with a local model.
        Document doc = new Document(getMyDir() + "Document.docx");

        AsposeLlmModel model = new AsposeLlmModel("Qwen25_3BPresetCpu");
        try /*JAVA: was using*/
        {
            Document translatedDoc = model.translate(doc, Language.GERMAN);
            translatedDoc.save(getArtifactsDir() + "AI.AsposeLlmTranslate.docx");
        }
        finally { if (model != null) model.dispose(); }
        //ExEnd:AsposeLlmTranslate
    }
}

