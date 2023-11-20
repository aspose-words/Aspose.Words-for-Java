// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Merger;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.MergeFormatMode;
import com.aspose.words.SaveFormat;
import com.aspose.words.Document;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import com.aspose.ms.System.IO.FileAccess;


@Test
class ExLowCode !Test class should be public in Java to run, please fix .Net source!  extends ApiExampleBase
{
    @Test
    public void mergeDocument() throws Exception
    {
        //ExStart
        //ExFor:Merger.Merge(String, String[])
        //ExFor:Merger.Merge(String[], MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
        //ExSummary:Shows how to merge documents into a single output document.
        //There is a several ways to merge documents:
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SimpleMerge.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" });

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveOptions.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, new OoxmlSaveOptions(); { .setPassword("Aspose.Words"); }, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveFormat.pdf", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        Document doc = Merger.merge(new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.DocumentInstance.docx");
        //ExEnd
    }

    @Test
    public void mergeStreamDocument() throws Exception
    {
        //ExStart            
        //ExFor:Merger.Merge(Stream[], MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
        //ExSummary:Shows how to merge documents from stream into a single output document.
        //There is a several ways to merge documents from stream:
        FileStream firstStreamIn = new FileStream(getMyDir() + "Big document.docx", FileMode.OPEN, FileAccess.READ);
        try /*JAVA: was using*/
        {
            FileStream secondStreamIn = new FileStream(getMyDir() + "Tables.docx", FileMode.OPEN, FileAccess.READ);
            try /*JAVA: was using*/
            {
                FileStream streamOut = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.SaveOptions.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{
                    Merger.mergeInternal(streamOut, new FileStream[] { firstStreamIn, secondStreamIn }, new OoxmlSaveOptions(); { .setPassword("Aspose.Words"); }, MergeFormatMode.KEEP_SOURCE_FORMATTING);
            	}
                finally { if (streamOut != null) streamOut.close(); }

                FileStream streamOut1 = new FileStream(getArtifactsDir() + "LowCode.MergeStreamDocument.SaveFormat.docx", FileMode.CREATE, FileAccess.READ_WRITE);
                try /*JAVA: was using*/
            	{                    
                    Merger.mergeInternal(streamOut1, new FileStream[] { firstStreamIn, secondStreamIn }, SaveFormat.DOCX);
            	}
                finally { if (streamOut1 != null) streamOut1.close(); }
               
                Document doc = Merger.mergeInternal(new FileStream[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MERGE_FORMATTING);
                doc.save(getArtifactsDir() + "LowCode.MergeStreamDocument.DocumentInstance.docx");
            }
            finally { if (secondStreamIn != null) secondStreamIn.close(); }
        }
        finally { if (firstStreamIn != null) firstStreamIn.close(); }
        //ExEnd
    }
}

