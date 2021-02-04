// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
import com.aspose.ms.System.Guid;
import com.aspose.words.BuildingBlockType;
import com.aspose.words.BuildingBlockGallery;
import com.aspose.words.BuildingBlockBehavior;
import com.aspose.words.DocumentVisitor;
import com.aspose.words.VisitorAction;
import com.aspose.words.Section;
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
    public void createAndInsert() throws Exception
    {
        // A document's glossary document stores building blocks.
        Document doc = new Document();
        GlossaryDocument glossaryDoc = new GlossaryDocument();
        doc.setGlossaryDocument(glossaryDoc);

        // Create a building block, name it, and then add it to the glossary document.
        BuildingBlock block = new BuildingBlock(glossaryDoc);
        {
            block.setName("Custom Block");
        }

        glossaryDoc.appendChild(block);

        // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
        Assert.assertEquals("00000000-0000-0000-0000-000000000000", block.getGuidInternal().toString());

        block.setGuidInternal(Guid.newGuid());

        // The following properties categorize building blocks
        // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
        Assert.assertEquals("(Empty Category)", block.getCategory());
        Assert.assertEquals(BuildingBlockType.NONE, block.getType());
        Assert.assertEquals(BuildingBlockGallery.ALL, block.getGallery());
        Assert.assertEquals(BuildingBlockBehavior.CONTENT, block.getBehavior());

        // Before we can add this building block to our document, we will need to give it some contents,
        // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
        BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
        block.accept(visitor);

        // We can access the block that we just made from the glossary document.
        BuildingBlock customBlock = glossaryDoc.getBuildingBlock(BuildingBlockGallery.QUICK_PARTS,
            "My custom building blocks", "Custom Block");

        // The block itself is a section that contains the text.
        Assert.assertEquals($"Text inside {customBlock.Name}\f", customBlock.getFirstSection().getBody().getFirstParagraph().getText());
        Assert.assertEquals(customBlock.getFirstSection(), customBlock.getLastSection());
        Assert.DoesNotThrow(() => Guid.parse(customBlock.getGuidInternal().toString())); //ExSkip
        Assert.assertEquals("My custom building blocks", customBlock.getCategory()); //ExSkip
        Assert.assertEquals(BuildingBlockType.NONE, customBlock.getType()); //ExSkip
        Assert.assertEquals(BuildingBlockGallery.QUICK_PARTS, customBlock.getGallery()); //ExSkip
        Assert.assertEquals(BuildingBlockBehavior.PARAGRAPH, customBlock.getBehavior()); //ExSkip

        // Now, we can insert it into the document as a new section.
        doc.appendChild(doc.importNode(customBlock.getFirstSection(), true));

        // We can also find it in Microsoft Word's Building Blocks Organizer and place it manually.
        doc.save(getArtifactsDir() + "BuildingBlocks.CreateAndInsert.dotx");
    }

    /// <summary>
    /// Sets up a visited building block to be inserted into the document as a quick part and adds text to its contents.
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
            // Configure the building block as a quick part, and add properties used by Building Blocks Organizer.
            block.setBehavior(BuildingBlockBehavior.PARAGRAPH);
            block.setCategory("My custom building blocks");
            block.setDescription("Using this block in the Quick Parts section of word will place its contents at the cursor.");
            block.setGallery(BuildingBlockGallery.QUICK_PARTS);

            // Add a section with text.
            // Inserting the block into the document will append this section with its child nodes at the location.
            Section section = new Section(mGlossaryDoc);
            block.appendChild(section);
            block.getFirstSection().ensureMinimum();

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
    //ExSummary:Shows ways of accessing building blocks in a glossary document.
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

        // There are various ways of accessing building blocks.
        // 1 -  Get the first/last building blocks in the collection:
        Assert.assertEquals("Block 1", glossaryDoc.getFirstBuildingBlock().getName());
        Assert.assertEquals("Block 5", glossaryDoc.getLastBuildingBlock().getName());

        // 2 -  Get a building block by index:
        Assert.assertEquals("Block 2", glossaryDoc.getBuildingBlocks().get(1).getName());
        Assert.assertEquals("Block 3", glossaryDoc.getBuildingBlocks().toArray()[2].getName());

        // 3 -  Get the first building block that matches a gallery, name and category:
        Assert.assertEquals("Block 4", 
            glossaryDoc.getBuildingBlock(BuildingBlockGallery.ALL, "(Empty Category)", "Block 4").getName());

        // We will do that using a custom visitor,
        // which will give every BuildingBlock in the GlossaryDocument a unique GUID
        GlossaryDocVisitor visitor = new GlossaryDocVisitor();
        glossaryDoc.accept(visitor);
        Assert.assertEquals(5, visitor.getDictionary().size()); //ExSkip

        System.out.println(visitor.getText());

        // In Microsoft Word, we can access the building blocks via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
        doc.save(getArtifactsDir() + "BuildingBlocks.GlossaryDocument.dotx"); 
    }

    /// <summary>
    /// Gives each building block in a visited glossary document a unique GUID.
    /// Stores the GUID-building block pairs in a dictionary.
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
