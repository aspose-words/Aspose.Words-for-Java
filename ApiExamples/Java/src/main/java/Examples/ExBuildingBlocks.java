package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.text.MessageFormat;
import java.util.HashMap;
import java.util.UUID;

@Test
public class ExBuildingBlocks extends ApiExampleBase {
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
    public void buildingBlockFields() throws Exception {
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
        Assert.assertEquals(glossaryDoc.getCount(), 1);

        // All GUIDs are this value by default
        Assert.assertEquals(block.getGuid().toString(), "00000000-0000-0000-0000-000000000000");

        // In Microsoft Word, we can use these attributes to find blocks in Insert > Quick Parts > Building Blocks Organizer
        Assert.assertEquals(block.getCategory(), "(Empty Category)");
        Assert.assertEquals(block.getType(), BuildingBlockType.NONE);
        Assert.assertEquals(block.getGallery(), BuildingBlockGallery.ALL);
        Assert.assertEquals(block.getBehavior(), BuildingBlockBehavior.CONTENT);

        // If we want to use our building block as an AutoText quick part, we need to give it some text and change some properties
        // All the necessary preparation will be done in a custom document visitor that we will accept
        BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
        block.accept(visitor);

        // We can find the block we made in the glossary document like this
        BuildingBlock customBlock = glossaryDoc.getBuildingBlock(BuildingBlockGallery.QUICK_PARTS,
                "My custom building blocks", "Custom Block");

        // Our block contains one section which now contains our text
        Assert.assertEquals(MessageFormat.format("Text inside {0}\f", customBlock.getName()), customBlock.getFirstSection().getBody().getFirstParagraph().getText());
        Assert.assertEquals(customBlock.getFirstSection(), customBlock.getLastSection());
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
    public static class BuildingBlockVisitor extends DocumentVisitor {
        public BuildingBlockVisitor(final GlossaryDocument ownerGlossaryDoc) {
            mBuilder = new StringBuilder();
            mGlossaryDoc = ownerGlossaryDoc;
        }

        public int visitBuildingBlockStart(final BuildingBlock block) {
            // Change values by default of created BuildingBlock
            block.setBehavior(BuildingBlockBehavior.PARAGRAPH);
            block.setCategory("My custom building blocks");
            block.setDescription("Using this block in the Quick Parts section of word will place its contents at the cursor.");
            block.setGallery(BuildingBlockGallery.QUICK_PARTS);

            block.setGuid(UUID.randomUUID());

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

        public int visitBuildingBlockEnd(final BuildingBlock block) {
            mBuilder.append("Visited " + block.getName() + "\r\n");
            return VisitorAction.CONTINUE;
        }

        private StringBuilder mBuilder;
        private GlossaryDocument mGlossaryDoc;
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
    public void glossaryDocument() throws Exception {
        Document doc = new Document();

        GlossaryDocument glossaryDoc = new GlossaryDocument();
        glossaryDoc.appendChild(createNewBuildingBlock(glossaryDoc, "Block 1"));
        glossaryDoc.appendChild(createNewBuildingBlock(glossaryDoc, "Block 2"));
        glossaryDoc.appendChild(createNewBuildingBlock(glossaryDoc, "Block 3"));
        glossaryDoc.appendChild(createNewBuildingBlock(glossaryDoc, "Block 4"));
        glossaryDoc.appendChild(createNewBuildingBlock(glossaryDoc, "Block 5"));

        Assert.assertEquals(glossaryDoc.getBuildingBlocks().getCount(), 5);

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

    public static BuildingBlock createNewBuildingBlock(final GlossaryDocument glossaryDoc, final String buildingBlockName) {
        BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
        buildingBlock.setName(buildingBlockName);

        return buildingBlock;
    }

    /// <summary>
    /// Simple implementation of giving each building block in a glossary document a unique GUID. Implemented as a Visitor.
    /// </summary>
    public static class GlossaryDocVisitor extends DocumentVisitor {
        public GlossaryDocVisitor() {
            mBlocksByGuid = new HashMap<>();
            mBuilder = new StringBuilder();
        }

        public String getText() {
            return mBuilder.toString();
        }

        public HashMap<UUID, BuildingBlock> getDictionary() {
            return mBlocksByGuid;
        }

        public int visitGlossaryDocumentStart(final GlossaryDocument glossary) {
            mBuilder.append("Glossary document found!\n");
            return VisitorAction.CONTINUE;
        }

        public int visitGlossaryDocumentEnd(final GlossaryDocument glossary) {
            mBuilder.append("Reached end of glossary!\n");
            mBuilder.append("BuildingBlocks found: " + mBlocksByGuid.size() + "\r\n");
            return VisitorAction.CONTINUE;
        }

        public int visitBuildingBlockStart(final BuildingBlock block) {
            Assert.assertEquals("00000000-0000-0000-0000-000000000000", block.getGuid().toString()); //ExSkip
            block.setGuid(UUID.randomUUID());
            mBlocksByGuid.put(block.getGuid(), block);
            return VisitorAction.CONTINUE;
        }

        public int visitBuildingBlockEnd(final BuildingBlock block) {
            mBuilder.append("\tVisited block \"" + block.getName() + "\"" + "\r\n");
            mBuilder.append("\t Type: " + block.getType() + "\r\n");
            mBuilder.append("\t Gallery: " + block.getGallery() + "\r\n");
            mBuilder.append("\t Behavior: " + block.getBehavior() + "\r\n");
            mBuilder.append("\t Description: " + block.getDescription() + "\r\n");

            return VisitorAction.CONTINUE;
        }

        private HashMap<UUID, BuildingBlock> mBlocksByGuid;
        private StringBuilder mBuilder;
    }
    //ExEnd
}
