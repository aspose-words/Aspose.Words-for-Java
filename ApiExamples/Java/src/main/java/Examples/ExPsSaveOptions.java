package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.annotations.Test;

@Test
public class ExPsSaveOptions extends ApiExampleBase {
    @Test
    public void useBookFoldPrintingSettings() throws Exception {
        //ExStart
        //ExFor:PsSaveOptions
        //ExFor:PsSaveOptions.SaveFormat
        //ExFor:PsSaveOptions.UseBookFoldPrintingSettings
        //ExSummary:Shows how to create a bookfold in the PostScript format.
        Document doc = new Document(getMyDir() + "Paragraphs.docx");

        // Configure both page setup and PsSaveOptions to create a book fold
        for (Section s : (Iterable<Section>) doc.getSections()) {
            s.getPageSetup().setMultiplePages(MultiplePagesType.BOOK_FOLD_PRINTING);
        }

        PsSaveOptions saveOptions = new PsSaveOptions();
        {
            saveOptions.setSaveFormat(SaveFormat.PS);
            saveOptions.setUseBookFoldPrintingSettings(true);
        }

        // In order to make a booklet, we will need to print this document, stack the pages
        // in the order they come out of the printer and then fold down the middle
        doc.save(getArtifactsDir() + "PsSaveOptions.UseBookFoldPrintingSettings.ps", saveOptions);
        //ExEnd
    }
}
