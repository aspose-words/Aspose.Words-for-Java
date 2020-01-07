// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.BorderCollection;
import java.util.Iterator;
import com.aspose.words.Border;
import java.awt.Color;
import com.aspose.words.LineStyle;


@Test
public class ExBorderCollection extends ApiExampleBase
{
    @Test
    public void getBordersEnumerator() throws Exception
    {
        //ExStart
        //ExFor:BorderCollection.GetEnumerator
        //ExSummary:Shows how to enumerate all borders in a collection.
        Document doc = new Document(getMyDir() + "Borders.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);

        BorderCollection borders = builder.getParagraphFormat().getBorders();

        Iterator<Border> enumerator = borders.iterator();
        try /*JAVA: was using*/
        {
            while (enumerator.hasNext())
            {
                // Do something useful.
                Border b = enumerator.next();
                b.setColor(Color.RoyalBlue);
                b.setLineStyle(LineStyle.DOUBLE);
            }
        }
        finally { if (enumerator != null) enumerator.close(); }

        doc.save(getArtifactsDir() + "Border.ChangedColorBorder.doc");                                                    
        //ExEnd
    }

    @Test
    public void removeAllBorders() throws Exception
    {
        //ExStart
        //ExFor:BorderCollection.ClearFormatting
        //ExSummary:Shows how to remove all borders from a paragraph at once.
        Document doc = new Document(getMyDir() + "Borders.doc");
        DocumentBuilder builder = new DocumentBuilder(doc);
        BorderCollection borders = builder.getParagraphFormat().getBorders();

        borders.clearFormatting();
        //ExEnd
    }
}
