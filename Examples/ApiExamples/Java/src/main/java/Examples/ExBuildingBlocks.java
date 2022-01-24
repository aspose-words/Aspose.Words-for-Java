package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
    public void createAndInsert() throws Exception {
        // A document's glossary document stores building blocks.
        Document doc = new Document();
        GlossaryDocument glossaryDoc = new GlossaryDocument();
        doc.setGlossaryDocument(glossaryDoc);

        // Create a building block, name it, and then add it to the glossary document.
        BuildingBlock block = new BuildingBlock(glossaryDoc);
        block.setName("Custom Block");

        glossaryDoc.appendChild(block);

        // All new building block GUIDs have the same zero value by default, and we can give them a new unique value.
        Assert.assertEquals(block.getGuid().toString(), "00000000-0000-0000-0000-000000000000");

        block.setGuid(UUID.randomUUID());

        // The following properties categorize building blocks
        // in the menu we can access in Microsoft Word via "Insert" -> "Quick Parts" -> "Building Blocks Organizer".
        Assert.assertEquals(block.getCategory(), "(Empty Category)");
        Assert.assertEquals(block.getType(), BuildingBlockType.NONE);
        Assert.assertEquals(block.getGallery(), BuildingBlockGallery.ALL);
        Assert.assertEquals(block.getBehavior(), BuildingBlockBehavior.CONTENT);

        // Before we can add this building block to our document, we will need to give it some contents,
        // which we will do using a document visitor. This visitor will also set a category, gallery, and behavior.
        BuildingBlockVisitor visitor = new BuildingBlockVisitor(glossaryDoc);
        block.accept(visitor);

        // We can access the block that we just made from the glossary document.
        BuildingBlock customBlock = glossaryDoc.getBuildingBlock(BuildingBlockGallery.QUICK_PARTS,
                "My custom building blocks", "Custom Block");

        // The block itself is a section that contains the text.
        Assert.assertEquals(MessageFormat.format("Text inside {0}\f", customBlock.getName()), customBlock.getFirstSection().getBody().getFirstParagraph().getText());
        Assert.assertEquals(customBlock.getFirstSection(), customBlock.getLastSection());
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
    public static class BuildingBlockVisitor extends DocumentVisitor {
        public BuildingBlockVisitor(final GlossaryDocument ownerGlossaryDoc) {
            mBuilder = new StringBuilder();
            mGlossaryDoc = ownerGlossaryDoc;
        }

        public int visitBuildingBlockStart(final BuildingBlock block) {
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

        public int visitBuildingBlockEnd(final BuildingBlock block) {
            mBuilder.append("Visited " + block.getName() + "\r\n");
            return VisitorAction.CONTINUE;
        }

        private final StringBuilder mBuilder;
        private final GlossaryDocument mGlossaryDoc;
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

    public static BuildingBlock createNewBuildingBlock(final GlossaryDocument glossaryDoc, final String buildingBlockName) {
        BuildingBlock buildingBlock = new BuildingBlock(glossaryDoc);
        buildingBlock.setName(buildingBlockName);

        return buildingBlock;
    }

    /// <summary>
    /// Gives each building block in a visited glossary document a unique GUID.
    /// Stores the GUID-building block pairs in a dictionary.
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

        private final HashMap<UUID, BuildingBlock> mBlocksByGuid;
        private final StringBuilder mBuilder;
    }
    //ExEnd
}
