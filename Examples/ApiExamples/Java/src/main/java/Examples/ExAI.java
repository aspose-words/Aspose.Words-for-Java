package Examples;

// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
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
        //GistId:366eb64fd56dec3c2eaa40410e594182
        //ExFor:GoogleAiModel
        //ExFor:OpenAiModel
        //ExFor:IAiModelText
        //ExFor:IAiModelText.Summarize(Document, SummarizeOptions)
        //ExFor:IAiModelText.Summarize(Document[], SummarizeOptions)
        //ExFor:SummarizeOptions
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
        IAiModelText model = (IAiModelText) AiModel.create(AiModelType.GPT_4_O_MINI).withApiKey(apiKey);

        SummarizeOptions options = new SummarizeOptions();
        options.setSummaryLength(SummaryLength.SHORT);
        Document oneDocumentSummary = model.summarize(firstDoc, options);
        oneDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.One.docx");

        options = new SummarizeOptions();
        options.setSummaryLength(SummaryLength.LONG);
        Document multiDocumentSummary = model.summarize(new Document[] { firstDoc, secondDoc }, options);
        multiDocumentSummary.save(getArtifactsDir() + "AI.AiSummarize.Multi.docx");
        //ExEnd:AiSummarize
    }
}

