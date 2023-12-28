package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.XlsxSectionMode;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.XlsxSaveOptions;
import com.aspose.words.CompressionLevel;

@Test
class ExXlsxSaveOptions extends ApiExampleBase
{
    @Test
    public void compressXlsx() throws Exception
    {
        //ExStart
        //ExFor:XlsxSaveOptions.CompressionLevel
        //ExSummary:Shows how to compress XLSX document.
        Document doc = new Document(getMyDir() + "Shape with linked chart.docx");

        XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();
        xlsxSaveOptions.setCompressionLevel(CompressionLevel.MAXIMUM); 

        doc.save(getArtifactsDir() + "XlsxSaveOptions.CompressXlsx.xlsx", xlsxSaveOptions);
        //ExEnd
    }

    @Test
    public void selectionMode() throws Exception
    {
        //ExStart:SelectionMode
        //GistId:66dd22f0854357e394a013b536e2181b
        //ExFor:XlsxSaveOptions.SectionMode
        //ExSummary:Shows how to save document as a separate worksheets.
        Document doc = new Document(getMyDir() + "Big document.docx");

        // Each section of a document will be created as a separate worksheet.
        // Use 'SingleWorksheet' to display all document on one worksheet.
        XlsxSaveOptions xlsxSaveOptions = new XlsxSaveOptions();
        xlsxSaveOptions.setSectionMode(XlsxSectionMode.MULTIPLE_WORKSHEETS);

        doc.save(getArtifactsDir() + "XlsxSaveOptions.SelectionMode.xlsx", xlsxSaveOptions);
        //ExEnd:SelectionMode
    }
}
