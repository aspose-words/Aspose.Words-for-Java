package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.List;
import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;

public class ExCleanupOptions extends ApiExampleBase {
    @Test
    public void removeUnusedResources() throws Exception {
        //ExStart
        //ExFor:Document.Cleanup(CleanupOptions)
        //ExFor:CleanupOptions
        //ExFor:CleanupOptions.UnusedLists
        //ExFor:CleanupOptions.UnusedStyles
        //ExFor:CleanupOptions.UnusedBuiltinStyles
        //ExSummary:Shows how to remove all unused custom styles from a document. 
        Document doc = new Document();

        doc.getStyles().add(StyleType.LIST, "MyListStyle1");
        doc.getStyles().add(StyleType.LIST, "MyListStyle2");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle1");
        doc.getStyles().add(StyleType.CHARACTER, "MyParagraphStyle2");

        // Combined with the built-in styles, the document now has eight styles.
        // A custom style is marked as "used" while there is any text within the document
        // formatted in that style. This means that the 4 styles we added are currently unused.
        Assert.assertEquals(8, doc.getStyles().getCount());

        // Apply a custom character style, and then a custom list style. Doing so will mark them as "used".
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getFont().setStyle(doc.getStyles().get("MyParagraphStyle1"));
        builder.writeln("Hello world!");

        List list = doc.getLists().add(doc.getStyles().get("MyListStyle1"));
        builder.getListFormat().setList(list);
        builder.writeln("Item 1");
        builder.writeln("Item 2");

        // Now, there is one unused character style and one unused list style.
        // The Cleanup() method, when configured with a CleanupOptions object, can target unused styles and remove them.
        CleanupOptions cleanupOptions = new CleanupOptions();
        cleanupOptions.setUnusedLists(true);
        cleanupOptions.setUnusedStyles(true);
        cleanupOptions.setUnusedBuiltinStyles(true);

        doc.cleanup(cleanupOptions);

        Assert.assertEquals(4, doc.getStyles().getCount());

        // Removing every node that a custom style is applied to marks it as "unused" again. 
        // Rerun the Cleanup method to remove them.
        doc.getFirstSection().getBody().removeAllChildren();
        doc.cleanup(cleanupOptions);

        Assert.assertEquals(2, doc.getStyles().getCount());
        //ExEnd
    }

    @Test
    public void removeDuplicateStyles() throws Exception {
        //ExStart
        //ExFor:CleanupOptions.DuplicateStyle
        //ExSummary:Shows how to remove duplicated styles from the document.
        Document doc = new Document();

        // Add two styles to the document with identical properties,
        // but different names. The second style is considered a duplicate of the first.
        Style myStyle = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle1");
        myStyle.getFont().setSize(14.0);
        myStyle.getFont().setName("Courier New");
        myStyle.getFont().setColor(Color.BLUE);

        Style duplicateStyle = doc.getStyles().add(StyleType.PARAGRAPH, "MyStyle2");
        duplicateStyle.getFont().setSize(14.0);
        duplicateStyle.getFont().setName("Courier New");
        duplicateStyle.getFont().setColor(Color.BLUE);

        Assert.assertEquals(6, doc.getStyles().getCount());

        // Apply both styles to different paragraphs within the document.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.getParagraphFormat().setStyleName(myStyle.getName());
        builder.writeln("Hello world!");

        builder.getParagraphFormat().setStyleName(duplicateStyle.getName());
        builder.writeln("Hello again!");

        ParagraphCollection paragraphs = doc.getFirstSection().getBody().getParagraphs();

        Assert.assertEquals(myStyle, paragraphs.get(0).getParagraphFormat().getStyle());
        Assert.assertEquals(duplicateStyle, paragraphs.get(1).getParagraphFormat().getStyle());

        // Configure a CleanOptions object, then call the Cleanup method to substitute all duplicate styles
        // with the original and remove the duplicates from the document.
        CleanupOptions cleanupOptions = new CleanupOptions();
        cleanupOptions.setDuplicateStyle(true);

        doc.cleanup(cleanupOptions);

        Assert.assertEquals(5, doc.getStyles().getCount());
        Assert.assertEquals(myStyle, paragraphs.get(0).getParagraphFormat().getStyle());
        Assert.assertEquals(myStyle, paragraphs.get(1).getParagraphFormat().getStyle());
        //ExEnd
    }
}

