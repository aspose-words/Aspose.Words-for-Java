// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.words.IAiModelText;
import com.aspose.words.OpenAiModel;
import com.aspose.words.AiModel;
import com.aspose.words.AiModelType;
import com.aspose.words.SummarizeOptions;
import com.aspose.words.SummaryLength;
import com.aspose.words.Language;
import com.aspose.words.CheckGrammarOptions;


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
        //ExFor:IAiModelText
        //ExFor:IAiModelText.Summarize(Document, SummarizeOptions)
        //ExFor:IAiModelText.Summarize(Document[], SummarizeOptions)
        //ExFor:SummarizeOptions
        //ExFor:SummarizeOptions.#ctor
        //ExFor:SummarizeOptions.SummaryLength
        //ExFor:SummaryLength
        //ExFor:AiModel
        //ExFor:AiModel.Create(AiModelType)
        //ExFor:AiModel.WithApiKey(String)
        //ExFor:AiModelType
        //ExSummary:Shows how to summarize text using OpenAI and Google models.
        Document firstDoc = new Document(getMyDir() + "Big document.docx");
        Document secondDoc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI or Google generative language models.
        IAiModelText model = ((OpenAiModel)AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey)).withOrganization("Organization").withProject("Project");

        Document oneDocumentSummary = model.summarize(firstDoc, new SummarizeOptions(); { oneDocumentSummary.setSummaryLength(SummaryLength.SHORT); });
        oneDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.One.docx");

        Document multiDocumentSummary = model.summarize(new Document[] { firstDoc, secondDoc }, new SummarizeOptions(); { multiDocumentSummary.setSummaryLength(SummaryLength.LONG); });
        multiDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.Multi.docx");
        //ExEnd:AiSummarize
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiTranslate() throws Exception
    {
        //ExStart:AiTranslate
        //GistId:695136dbbe4f541a8a0a17b3d3468689
        //ExFor:IAiModelText.Translate(Document, AI.Language)
        //ExFor:AI.Language
        //ExSummary:Shows how to translate text using Google models.
        Document doc = new Document(getMyDir() + "Document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use Google generative language models.
        IAiModelText model = (IAiModelText)AiModel.create(AiModelType.GEMINI_15_FLASH).withApiKey(apiKey);

        Document translatedDoc = model.translate(doc, Language.ARABIC);
        translatedDoc.save(getArtifactsDir() + "AI.AiTranslate.docx");
        //ExEnd:AiTranslate
    }

    @Test (enabled = false, description = "This test should be run manually to manage API requests amount")
    public void aiGrammar() throws Exception
    {
        //ExStart:AiGrammar
        //GistId:f86d49dc0e6781b93e576539a01e6ca2
        //ExFor:IAiModelText.CheckGrammar(Document, CheckGrammarOptions)
        //ExFor:CheckGrammarOptions
        //ExSummary:Shows how to check the grammar of a document.
        Document doc = new Document(getMyDir() + "Big document.docx");

        String apiKey = System.getenv("API_KEY");
        // Use OpenAI generative language models.
        IAiModelText model = (IAiModelText)AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey);

        CheckGrammarOptions grammarOptions = new CheckGrammarOptions();
        grammarOptions.setImproveStylistics(true);

        Document proofedDoc = model.checkGrammar(doc, grammarOptions);
        proofedDoc.save("AI.AiGrammar.docx");
        //ExEnd:AiGrammar
    }
}

