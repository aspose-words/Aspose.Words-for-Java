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

import java.awt.*;
import java.util.Iterator;

public class ExBorderCollection extends ApiExampleBase {
    @Test
    public void getBordersEnumerator() throws Exception {
        //ExStart
        //ExFor:BorderCollection.GetEnumerator
        //ExSummary:Shows how to enumerate all borders in a collection.
        Document doc = new Document(getMyDir() + "Borders.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();

        Iterator<Border> enumerator = borders.iterator();
        while (enumerator.hasNext()) {
            // Do something useful.
            Border b = enumerator.next();
            b.setColor(new Color(65, 105, 225));//RoyalBlue
            b.setLineStyle(LineStyle.DOUBLE);
        }

        doc.save(getArtifactsDir() + "BorderCollection.GetBordersEnumerator.doc");
        //ExEnd
    }

    @Test
    public void removeAllBorders() throws Exception {
        //ExStart
        //ExFor:BorderCollection.ClearFormatting
        //ExSummary:Shows how to remove all borders from a paragraph at once.
        Document doc = new Document(getMyDir() + "Borders.docx");
        DocumentBuilder builder = new DocumentBuilder(doc);
        BorderCollection borders = builder.getParagraphFormat().getBorders();

        borders.clearFormatting();
        //ExEnd
    }
}
