package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
}
