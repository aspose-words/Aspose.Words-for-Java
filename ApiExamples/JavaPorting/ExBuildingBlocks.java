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
import com.aspose.words.GlossaryDocument;
import com.aspose.words.BuildingBlock;
import org.testng.Assert;
import com.aspose.words.BuildingBlockType;
import com.aspose.words.BuildingBlockGallery;
import com.aspose.words.BuildingBlockBehavior;
import com.aspose.ms.System.Guid;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.Section;
import com.aspose.words.Body;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.ms.System.Text.msStringBuilder;
import com.aspose.ms.System.msConsole;
import java.util.HashMap;
import com.aspose.ms.System.Collections.msDictionary;


@Test
public class ExBuildingBlocks extends ApiExampleBase
{
    //ExStart
    //ExFor:Document.GlossaryDocument
    //ExFor:BuildingBlocks.BuildingBlock
    //ExFor:BuildingBlocks.BuildingBlock.#ctor(BuildingBlocks.GlossaryDocument) 
    //ExFor:BuildingBlocks.BuildingBlock.Accept(DocumentVisitor)
    //ExFor:BuildingBlocks.BuildingBlock.Behavior
    //ExFor:BuildingBlocks.BuildingBlock.Category
    //ExFor:BuildingBlocks.BuildingBlock.Description
    //ExFor:BuildingBlocks.BuildingBlock.FirstSection
    //ExFor:BuildingBlocks.BuildingBlock.Gallery
    //ExFor:BuildingBlocks.BuildingBlock.Guid
    //ExFor:BuildingBlocks.BuildingBlock.LastSection
    //ExFor:BuildingBlocks.BuildingBlock.Name
    //ExFor:BuildingBlocks.BuildingBlock.Sections
    //ExFor:BuildingBlocks.BuildingBlock.Type
    //ExFor:BuildingBlocks.BuildingBlockBehavior
    //ExFor:BuildingBlocks.BuildingBlockType
    //ExSummary:Shows how to add a custom building block to a document.
    @Test //ExSkip
    public void buildingBlockFields() throws Exception
    {
        Document doc = new Document();

        // BuildingBlocks are stored inside the glossary document
        // If you're making a document from scratch, the glossary document must also be manually created
        GlossaryDocument glossaryDoc = new GlossaryDocument();
        doc.setGlossaryDocument(glossaryDoc);

        // Create a building block and name it
        BuildingBlock block = new BuildingBlock(glossaryDoc);
        block.setName("Custom Block");
        
        // Put in in the document's glossary document
        glossaryDoc.appendChild(block);
        Assert.assertEquals(1, glossaryDoc.getCount());

        // All GUIDs are this value by default
        Assert.assertEquals("00000000-0000-0000-0000-000000000000", block.getGuidInternal().toString());

        // In Microsoft Word, we can use these attributes to find blocks in Insert > Quick Parts > Building Blocks Organizer  
        Assert.assertEquals("(Empty Category)", block.getCategory());
        Assert.assertEquals(BuildingBlockType.NONE, block.getType());
        Assert.assertEquals(BuildingBlockGallery.ALL, block.getGallery());
        Assert.assertEquals(BuildingBlockBehavior.CONTENT, block.getBehavior());

        // If we want to use our building block as an AutoText quick part, we need to give it some text and change some properties
        // All the necessary preparation will be done in a custom document visitor that we will accept
        BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
        block.accept(visitor);

        // We can find the block we made in the glossary document like this
        BuildingBlock customBlock = glossaryDoc.getBuildingBlock(BuildingBlockGallery.QUICK_PARTS,
            "My custom building blocks", "Custom Block");

        // Our block contains one section which now contains our text
        Assert.assertEquals($"Text inside {customBlock.Name}\f", customBlock.getFirstSection().getBody().getFirstParagraph().getText());
        Assert.assertEquals(customBlock.getFirstSection(), customBlock.getLastSection());
        Assert.DoesNotThrow(() => Guid.parse(customBlock.getGuidInternal().toString())); //ExSkip
        Assert.assertEquals("My custom building blocks", customBlock.getCategory()); //ExSkip
        Assert.assertEquals(BuildingBlockType.NONE, customBlock.getType()); //ExSkip
        Assert.assertEquals(BuildingBlockGallery.QUICK_PARTS, customBlock.getGallery()); //ExSkip
        Assert.assertEquals(BuildingBlockBehavior.PARAGRAPH, customBlock.getBehavior()); //ExSkip

        // Then we can insert it into the document as a new section
        doc.appendChild(doc.importNode(customBlock.getFirstSection(), true));

        // Or we can find it in Microsoft Word's Building Blocks Organizer and place it manually
        doc.save(getArtifactsDir() + "BuildingBlocks.BuildingBlockFields.dotx");
    }

    /// <summary>
    /// Simple implementation of adding text to a building block and preparing it for usage in the document. Implemented as a Visitor.
    /// </summary>
    public static class BuildingBlockVisitor extends DocumentVisitor
    {
        public BuildingBlockVisitor(GlossaryDocument ownerGlossaryDoc)
        {
            mBuilder = new StringBuilder();
            mGlossaryDoc = ownerGlossaryDoc;
        }

        public /*override*/ /*VisitorAction*/int visitBuildingBlockStart(BuildingBlock block)
        {
            // Change values by default of created BuildingBlock
            block.setBehavior(BuildingBlockBehavior.PARAGRAPH);
            block.setCategory("My custom building blocks");
            block.setDescription("Using this block in the Quick Parts section of word will place its contents at the cursor.");
            block.setGallery(BuildingBlockGallery.QUICK_PARTS);

            block.setGuidInternal(Guid.newGuid());

            // Add content for the BuildingBlock to have an effect when used in the document
            Section section = new Section(mGlossaryDoc);
            block.appendChild(section);

            Body body = new Body(mGlossaryDoc);
            section.appendChild(body);

            Paragraph paragraph = new Paragraph(mGlossaryDoc);
            body.appendChild(paragraph);

            // Add text that will be visible in the document
            Run run = new Run(mGlossaryDoc, "Text inside " + block.getName());
            block.getFirstSection().getBody().getFirstParagraph().appendChild(run);

            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBuildingBlockEnd(BuildingBlock block)
        {
            msStringBuilder.append(mBuilder, "Visited " + block.getName() + "\r\n");
            return VisitorAction.CONTINUE;
        }

        private /*final*/ StringBuilder mBuilder;
        private /*final*/ GlossaryDocument mGlossaryDoc;
    }
    //ExEnd

    //ExStart
    //ExFor:BuildingBlocks.GlossaryDocument
    //ExFor:BuildingBlocks.GlossaryDocument.Accept(DocumentVisitor)
    //ExFor:BuildingBlocks.GlossaryDocument.BuildingBlocks
    //ExFor:BuildingBlocks.GlossaryDocument.FirstBuildingBlock
    //ExFor:BuildingBlocks.GlossaryDocument.GetBuildingBlock(BuildingBlocks.BuildingBlockGallery,System.String,System.String)
    //ExFor:BuildingBlocks.GlossaryDocument.LastBuildingBlock
    //ExFor:BuildingBlocks.BuildingBlockCollection
    //ExFor:BuildingBlocks.BuildingBlockCollection.Item(System.Int32)
    //ExFor:BuildingBlocks.BuildingBlockCollection.ToArray
    //ExFor:BuildingBlocks.BuildingBlockGallery
    //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlock)
    //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlock)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(GlossaryDocument)
    //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(GlossaryDocument)
    //ExSummary:Shows how to use GlossaryDocument and BuildingBlockCollection.
    @Test //ExSkip
    public void glossaryDocument() throws Exception
    {
        Document doc = new Document();

        GlossaryDocument glossaryDoc = new GlossaryDocument();
        glossaryDoc.appendChild(new BuildingBlock(glossaryDoc); { .setName("Block 1"); });
        glossaryDoc.appendChild(new BuildingBlock(glossaryDoc); { .setName("Block 2"); });
        glossaryDoc.appendChild(new BuildingBlock(glossaryDoc); { .setName("Block 3"); });
        glossaryDoc.appendChild(new BuildingBlock(glossaryDoc); { .setName("Block 4"); });
        glossaryDoc.appendChild(new BuildingBlock(glossaryDoc); { .setName("Block 5"); });

        Assert.assertEquals(5, glossaryDoc.getBuildingBlocks().getCount());

        doc.setGlossaryDocument(glossaryDoc);

        // There is a different ways how to get created building blocks
        Assert.assertEquals("Block 1", glossaryDoc.getFirstBuildingBlock().getName());
        Assert.assertEquals("Block 2", glossaryDoc.getBuildingBlocks().get(1).getName());
        Assert.assertEquals("Block 3", glossaryDoc.getBuildingBlocks().toArray()[2].getName());
        Assert.assertEquals("Block 4", glossaryDoc.getBuildingBlock(BuildingBlockGallery.ALL, "(Empty Category)", "Block 4").getName());
        Assert.assertEquals("Block 5", glossaryDoc.getLastBuildingBlock().getName());

        // We will do that using a custom visitor, which also will give every BuildingBlock in the GlossaryDocument a unique GUID
        GlossaryDocVisitor visitor = new GlossaryDocVisitor();
        glossaryDoc.accept(visitor);
        Assert.assertEquals(5, visitor.getDictionary().size()); //ExSkip

        System.out.println(visitor.getText());

        // We can find our new blocks in Microsoft Word via Insert > Quick Parts > Building Blocks Organizer...
        doc.save(getArtifactsDir() + "BuildingBlocks.GlossaryDocument.dotx"); 
    }

    /// <summary>
    /// Simple implementation of giving each building block in a glossary document a unique GUID. Implemented as a Visitor.
    /// </summary>
    public static class GlossaryDocVisitor extends DocumentVisitor
    {
        public GlossaryDocVisitor()
        {
            mBlocksByGuid = new HashMap<Guid, BuildingBlock>();
            mBuilder = new StringBuilder();
        }

        public String getText()
        {
            return mBuilder.toString();
        }

        public HashMap<Guid, BuildingBlock> getDictionary()
        {
            return mBlocksByGuid;
        }

        public /*override*/ /*VisitorAction*/int visitGlossaryDocumentStart(GlossaryDocument glossary)
        {
            msStringBuilder.appendLine(mBuilder, "Glossary document found!");
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitGlossaryDocumentEnd(GlossaryDocument glossary)
        {
            msStringBuilder.appendLine(mBuilder, "Reached end of glossary!");
            msStringBuilder.appendLine(mBuilder, "BuildingBlocks found: " + mBlocksByGuid.size());
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBuildingBlockStart(BuildingBlock block)
        {
            Assert.assertEquals("00000000-0000-0000-0000-000000000000", block.getGuidInternal().toString()); //ExSkip
            block.setGuidInternal(Guid.newGuid());
            msDictionary.add(mBlocksByGuid, block.getGuidInternal(), block);
            return VisitorAction.CONTINUE;
        }

        public /*override*/ /*VisitorAction*/int visitBuildingBlockEnd(BuildingBlock block)
        {
            msStringBuilder.appendLine(mBuilder, "\tVisited block \"" + block.getName() + "\"");
            msStringBuilder.appendLine(mBuilder, "\t Type: " + block.getType());
            msStringBuilder.appendLine(mBuilder, "\t Gallery: " + block.getGallery());
            msStringBuilder.appendLine(mBuilder, "\t Behavior: " + block.getBehavior());
            msStringBuilder.appendLine(mBuilder, "\t Description: " + block.getDescription());

            return VisitorAction.CONTINUE;
        }

        private /*final*/ HashMap<Guid, BuildingBlock> mBlocksByGuid;
        private /*final*/ StringBuilder mBuilder;
    }
    //ExEnd
}
