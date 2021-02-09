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

import java.awt.*;

public class ExSection extends ApiExampleBase {
    @Test
    public void protect() throws Exception {
        //ExStart
        //ExFor:Document.Protect(ProtectionType)
        //ExFor:ProtectionType
        //ExFor:Section.ProtectedForForms
        //ExSummary:Shows how to turn off protection for a section.
        Document doc = new Document();

        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Section 1. Hello world!");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        builder.writeln("Section 2. Hello again!");
        builder.write("Please enter text here: ");
        builder.insertTextInput("TextInput1", TextFormFieldType.REGULAR, "", "Placeholder text", 0);

        // Apply write protection to every section in the document.
        doc.protect(ProtectionType.ALLOW_ONLY_FORM_FIELDS);

        // Turn off write protection for the first section.
        doc.getSections().get(0).setProtectedForForms(false);

        // In this output document, we will be able to edit the first section freely,
        // and we will only be able to edit the contents of the form field in the second section.
        doc.save(getArtifactsDir() + "Section.Protect.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Section.Protect.docx");

        Assert.assertFalse(doc.getSections().get(0).getProtectedForForms());
        Assert.assertTrue(doc.getSections().get(1).getProtectedForForms());
    }

    @Test
    public void addRemove() throws Exception {
        //ExStart
        //ExFor:Document.Sections
        //ExFor:Section.Clone
        //ExFor:SectionCollection
        //ExFor:NodeCollection.RemoveAt(Int32)
        //ExSummary:Shows how to add and remove sections in a document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        Assert.assertEquals("Section 1\fSection 2", doc.getText().trim());

        // Delete the first section from the document.
        doc.getSections().removeAt(0);

        Assert.assertEquals("Section 2", doc.getText().trim());

        // Append a copy of what is now the first section to the end of the document.
        int lastSectionIdx = doc.getSections().getCount() - 1;
        Section newSection = doc.getSections().get(lastSectionIdx).deepClone();
        doc.getSections().add(newSection);

        Assert.assertEquals("Section 2\fSection 2", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void firstAndLast() throws Exception {
        //ExStart
        //ExFor:Document.FirstSection
        //ExFor:Document.LastSection
        //ExSummary:Shows how to create a new section with a document builder.
        Document doc = new Document();

        // A blank document contains one section by default,
        // which contains child nodes that we can edit.
        Assert.assertEquals(1, doc.getSections().getCount());

        // Use a document builder to add text to the first section.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a second section by inserting a section break.
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);

        Assert.assertEquals(2, doc.getSections().getCount());

        // Each section has its own page setup settings.
        // We can split the text in the second section into two columns.
        // This will not affect the text in the first section.
        doc.getLastSection().getPageSetup().getTextColumns().setCount(2);
        builder.writeln("Column 1.");
        builder.insertBreak(BreakType.COLUMN_BREAK);
        builder.writeln("Column 2.");

        Assert.assertEquals(1, doc.getFirstSection().getPageSetup().getTextColumns().getCount());
        Assert.assertEquals(2, doc.getLastSection().getPageSetup().getTextColumns().getCount());

        doc.save(getArtifactsDir() + "Section.Create.docx");
        //ExEnd

        doc = new Document(getArtifactsDir() + "Section.Create.docx");

        Assert.assertEquals(1, doc.getFirstSection().getPageSetup().getTextColumns().getCount());
        Assert.assertEquals(2, doc.getLastSection().getPageSetup().getTextColumns().getCount());
    }

    @Test
    public void createManually() throws Exception {
        //ExStart
        //ExFor:Node.GetText
        //ExFor:CompositeNode.RemoveAllChildren
        //ExFor:CompositeNode.AppendChild
        //ExFor:Section
        //ExFor:Section.#ctor
        //ExFor:Section.PageSetup
        //ExFor:PageSetup.SectionStart
        //ExFor:PageSetup.PaperSize
        //ExFor:SectionStart
        //ExFor:PaperSize
        //ExFor:Body
        //ExFor:Body.#ctor
        //ExFor:Paragraph
        //ExFor:Paragraph.#ctor
        //ExFor:Paragraph.ParagraphFormat
        //ExFor:ParagraphFormat
        //ExFor:ParagraphFormat.StyleName
        //ExFor:ParagraphFormat.Alignment
        //ExFor:ParagraphAlignment
        //ExFor:Run
        //ExFor:Run.#ctor(DocumentBase)
        //ExFor:Run.Text
        //ExFor:Inline.Font
        //ExSummary:Shows how to construct an Aspose.Words document by hand.
        Document doc = new Document();

        // A blank document contains one section, one body and one paragraph.
        // Call the "RemoveAllChildren" method to remove all those nodes,
        // and end up with a document node with no children.
        doc.removeAllChildren();

        // This document now has no composite child nodes that we can add content to.
        // If we wish to edit it, we will need to repopulate its node collection.
        // First, create a new section, and then append it as a child to the root document node.
        Section section = new Section(doc);
        doc.appendChild(section);

        // Set some page setup properties for the section.
        section.getPageSetup().setSectionStart(SectionStart.NEW_PAGE);
        section.getPageSetup().setPaperSize(PaperSize.LETTER);

        // A section needs a body, which will contain and display all its contents
        // on the page between the section's header and footer.
        Body body = new Body(doc);
        section.appendChild(body);

        // Create a paragraph, set some formatting properties, and then append it as a child to the body.
        Paragraph para = new Paragraph(doc);

        para.getParagraphFormat().setStyleName("Heading 1");
        para.getParagraphFormat().setAlignment(ParagraphAlignment.CENTER);

        body.appendChild(para);

        // Finally, add some content to do the document. Create a run,
        // set its appearance and contents, and then append it as a child to the paragraph.
        Run run = new Run(doc);
        run.setText("Hello World!");
        run.getFont().setColor(Color.RED);
        para.appendChild(run);

        Assert.assertEquals("Hello World!", doc.getText().trim());

        doc.save(getArtifactsDir() + "Section.CreateManually.docx");
        //ExEnd
    }

    @Test
    public void ensureMinimum() throws Exception {
        //ExStart
        //ExFor:NodeCollection.Add
        //ExFor:Section.EnsureMinimum
        //ExFor:SectionCollection.Item(Int32)
        //ExSummary:Shows how to prepare a new section node for editing.
        Document doc = new Document();

        // A blank document comes with a section, which has a body, which in turn has a paragraph.
        // We can add contents to this document by adding elements such as text runs, shapes, or tables to that paragraph.
        Assert.assertEquals(NodeType.SECTION, doc.getChild(NodeType.ANY, 0, true).getNodeType());
        Assert.assertEquals(NodeType.BODY, doc.getSections().get(0).getChild(NodeType.ANY, 0, true).getNodeType());
        Assert.assertEquals(NodeType.PARAGRAPH, doc.getSections().get(0).getBody().getChild(NodeType.ANY, 0, true).getNodeType());

        // If we add a new section like this, it will not have a body, or any other child nodes.
        doc.getSections().add(new Section(doc));

        Assert.assertEquals(0, doc.getSections().get(1).getChildNodes(NodeType.ANY, true).getCount());

        // Run the "EnsureMinimum" method to add a body and a paragraph to this section to begin editing it.
        doc.getLastSection().ensureMinimum();

        Assert.assertEquals(NodeType.BODY, doc.getSections().get(1).getChild(NodeType.ANY, 0, true).getNodeType());
        Assert.assertEquals(NodeType.PARAGRAPH, doc.getSections().get(1).getBody().getChild(NodeType.ANY, 0, true).getNodeType());

        doc.getSections().get(0).getBody().getFirstParagraph().appendChild(new Run(doc, "Hello world!"));

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void bodyEnsureMinimum() throws Exception {
        //ExStart
        //ExFor:Section.Body
        //ExFor:Body.EnsureMinimum
        //ExSummary:Clears main text from all sections from the document leaving the sections themselves.
        Document doc = new Document();

        // A blank document contains one section, one body and one paragraph.
        // Call the "RemoveAllChildren" method to remove all those nodes,
        // and end up with a document node with no children.
        doc.removeAllChildren();

        // This document now has no composite child nodes that we can add content to.
        // If we wish to edit it, we will need to repopulate its node collection.
        // First, create a new section, and then append it as a child to the root document node.
        Section section = new Section(doc);
        doc.appendChild(section);

        // A section needs a body, which will contain and display all its contents
        // on the page between the section's header and footer.
        Body body = new Body(doc);
        section.appendChild(body);

        // This body has no children, so we cannot add runs to it yet.
        Assert.assertEquals(0, doc.getFirstSection().getBody().getChildNodes(NodeType.ANY, true).getCount());

        // Call the "EnsureMinimum" to make sure that this body contains at least one empty paragraph. 
        body.ensureMinimum();

        // Now, we can add runs to the body, and get the document to display them.
        body.getFirstParagraph().appendChild(new Run(doc, "Hello world!"));

        Assert.assertEquals("Hello world!", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void bodyChildNodes() throws Exception {
        //ExStart
        //ExFor:Body.NodeType
        //ExFor:HeaderFooter.NodeType
        //ExFor:Document.FirstSection
        //ExSummary:Shows how to iterate through the children of a composite node.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.write("Primary header");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.write("Primary footer");

        Section section = doc.getFirstSection();

        // A Section is a composite node and can contain child nodes,
        // but only if those child nodes are of a "Body" or "HeaderFooter" node type.
        for (Node node : section) {
            switch (node.getNodeType()) {
                case NodeType.BODY: {
                    Body body = (Body) node;

                    System.out.println("Body:");
                    System.out.println("\t\"{body.GetText().Trim()}\"");
                    break;
                }
                case NodeType.HEADER_FOOTER: {
                    HeaderFooter headerFooter = (HeaderFooter) node;

                    System.out.println("HeaderFooter type: {headerFooter.HeaderFooterType}:");
                    System.out.println("\t\"{headerFooter.GetText().Trim()}\"");
                    break;
                }
                default: {
                    throw new Exception("Unexpected node type in a section.");
                }
            }
        }
        //ExEnd
    }

    @Test
    public void clear() throws Exception {
        //ExStart
        //ExFor:NodeCollection.Clear
        //ExSummary:Shows how to remove all sections from a document.
        Document doc = new Document(getMyDir() + "Document.docx");

        // This document has one section with a few child nodes containing and displaying all the document's contents.
        Assert.assertEquals(1, doc.getSections().getCount());
        Assert.assertEquals(19, doc.getSections().get(0).getChildNodes(NodeType.ANY, true).getCount());
        Assert.assertEquals("Hello World!\r\rHello Word!\r\r\rHello World!", doc.getText().trim());

        // Clear the collection of sections, which will remove all of the document's children.
        doc.getSections().clear();

        Assert.assertEquals(0, doc.getChildNodes(NodeType.ANY, true).getCount());
        Assert.assertEquals("", doc.getText().trim());
        //ExEnd
    }

    @Test
    public void prependAppendContent() throws Exception {
        //ExStart
        //ExFor:Section.AppendContent
        //ExFor:Section.PrependContent
        //ExSummary:Shows how to append the contents of a section to another section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 3");

        Section section = doc.getSections().get(2);

        Assert.assertEquals("Section 3" + ControlChar.SECTION_BREAK, section.getText());

        // Insert the contents of the first section to the beginning of the third section.
        Section sectionToPrepend = doc.getSections().get(0);
        section.prependContent(sectionToPrepend);

        // Insert the contents of the second section to the end of the third section.
        Section sectionToAppend = doc.getSections().get(1);
        section.appendContent(sectionToAppend);

        // The "PrependContent" and "AppendContent" methods did not create any new sections.
        Assert.assertEquals(3, doc.getSections().getCount());
        Assert.assertEquals("Section 1" + ControlChar.PARAGRAPH_BREAK +
                "Section 3" + ControlChar.PARAGRAPH_BREAK +
                "Section 2" + ControlChar.SECTION_BREAK, section.getText());
        //ExEnd
    }

    @Test
    public void clearContent() throws Exception {
        //ExStart
        //ExFor:Section.ClearContent
        //ExSummary:Shows how to clear the contents of a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Hello world!");

        Assert.assertEquals("Hello world!", doc.getText().trim());
        Assert.assertEquals(1, doc.getFirstSection().getBody().getParagraphs().getCount());

        // Running the "ClearContent" method will remove all the section contents
        // but leave a blank paragraph to add content again.
        doc.getFirstSection().clearContent();

        Assert.assertEquals("", doc.getText().trim());
        Assert.assertEquals(1, doc.getFirstSection().getBody().getParagraphs().getCount());
        //ExEnd
    }

    @Test
    public void clearHeadersFooters() throws Exception {
        //ExStart
        //ExFor:Section.ClearHeadersFooters
        //ExSummary:Shows how to clear the contents of all headers and footers in a section.
        Document doc = new Document();

        Assert.assertEquals(0, doc.getFirstSection().getHeadersFooters().getCount());

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.writeln("This is the primary header.");
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.writeln("This is the primary footer.");

        Assert.assertEquals(2, doc.getFirstSection().getHeadersFooters().getCount());

        Assert.assertEquals("This is the primary header.", doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getText().trim());
        Assert.assertEquals("This is the primary footer.", doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().trim());

        // Empty all the headers and footers in this section of all their contents.
        // The headers and footers themselves will still be present but will have nothing to display.
        doc.getFirstSection().clearHeadersFooters();

        Assert.assertEquals(2, doc.getFirstSection().getHeadersFooters().getCount());

        Assert.assertEquals("", doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getText().trim());
        Assert.assertEquals("", doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getText().trim());
        //ExEnd
    }

    @Test
    public void deleteHeaderFooterShapes() throws Exception {
        //ExStart
        //ExFor:Section.DeleteHeaderFooterShapes
        //ExSummary:Shows how to remove all shapes from all headers footers in a section.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a primary header with a shape.
        builder.moveToHeaderFooter(HeaderFooterType.HEADER_PRIMARY);
        builder.insertShape(ShapeType.RECTANGLE, 100.0, 100.0);

        // Create a primary footer with an image.
        builder.moveToHeaderFooter(HeaderFooterType.FOOTER_PRIMARY);
        builder.insertImage(getImageDir() + "Logo Icon.ico");

        Assert.assertEquals(1, doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getChildNodes(NodeType.SHAPE, true).getCount());
        Assert.assertEquals(1, doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getChildNodes(NodeType.SHAPE, true).getCount());

        // Remove all shapes from the headers and footers in the first section.
        doc.getFirstSection().deleteHeaderFooterShapes();

        Assert.assertEquals(0, doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.HEADER_PRIMARY).getChildNodes(NodeType.SHAPE, true).getCount());
        Assert.assertEquals(0, doc.getFirstSection().getHeadersFooters().getByHeaderFooterType(HeaderFooterType.FOOTER_PRIMARY).getChildNodes(NodeType.SHAPE, true).getCount());
        //ExEnd
    }

    @Test
    public void sectionsCloneSection() throws Exception {
        Document doc = new Document(getMyDir() + "Document.docx");
        Section cloneSection = doc.getSections().get(0).deepClone();
    }

    @Test
    public void sectionsImportSection() throws Exception {
        Document srcDoc = new Document(getMyDir() + "Document.docx");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section) dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
    }

    @Test
    public void migrateFrom2XImportSection() throws Exception {
        Document srcDoc = new Document();
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section) dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
    }

    @Test
    public void modifyPageSetupInAllSections() throws Exception {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.write("Section 1");
        builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        builder.write("Section 2");

        // It is important to understand that a document can contain many sections,
        // and each section has its page setup. In this case, we want to modify them all.
        for (Section section : doc.getSections())
            section.getPageSetup().setPaperSize(PaperSize.LETTER);

        doc.save(getArtifactsDir() + "Section.ModifyPageSetupInAllSections.doc");
    }

    @Test
    public void cultureInfoPageSetupDefaults() throws Exception {
        CurrentThreadSettings.setLocale("en-us");

        Document docEn = new Document();

        // Assert that page defaults comply with current culture info.
        Section sectionEn = docEn.getSections().get(0);
        Assert.assertEquals(sectionEn.getPageSetup().getLeftMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionEn.getPageSetup().getRightMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionEn.getPageSetup().getTopMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionEn.getPageSetup().getBottomMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionEn.getPageSetup().getHeaderDistance(), 36.0); // 1.27 cm
        Assert.assertEquals(sectionEn.getPageSetup().getFooterDistance(), 36.0); // 1.27 cm
        Assert.assertEquals(sectionEn.getPageSetup().getTextColumns().getSpacing(), 36.0); // 1.27 cm

        // Change the culture and assert that the page defaults are changed.
        CurrentThreadSettings.setLocale("de-de");

        Document docDe = new Document();

        Section sectionDe = docDe.getSections().get(0);
        Assert.assertEquals(sectionDe.getPageSetup().getLeftMargin(), 70.85); // 2.5 cm         
        Assert.assertEquals(sectionDe.getPageSetup().getRightMargin(), 70.85); // 2.5 cm
        Assert.assertEquals(sectionDe.getPageSetup().getTopMargin(), 70.85); // 2.5 cm
        Assert.assertEquals(sectionDe.getPageSetup().getBottomMargin(), 56.7); // 2 cm
        Assert.assertEquals(sectionDe.getPageSetup().getHeaderDistance(), 35.4); // 1.25 cm
        Assert.assertEquals(sectionDe.getPageSetup().getFooterDistance(), 35.4); // 1.25 cm
        Assert.assertEquals(sectionDe.getPageSetup().getTextColumns().getSpacing(), 35.4); // 1.25 cm

        // Change page defaults.
        sectionDe.getPageSetup().setLeftMargin(90.0); // 3.17 cm
        sectionDe.getPageSetup().setRightMargin(90.0); // 3.17 cm
        sectionDe.getPageSetup().setTopMargin(72.0); // 2.54 cm
        sectionDe.getPageSetup().setBottomMargin(72.0); // 2.54 cm
        sectionDe.getPageSetup().setHeaderDistance(35.4); // 1.25 cm
        sectionDe.getPageSetup().setFooterDistance(35.4); // 1.25 cm
        sectionDe.getPageSetup().getTextColumns().setSpacing(35.4); // 1.25 cm

        docDe = DocumentHelper.saveOpen(docDe);

        Section sectionDeAfter = docDe.getSections().get(0);
        Assert.assertEquals(sectionDeAfter.getPageSetup().getLeftMargin(), 90.0); // 3.17 cm         
        Assert.assertEquals(sectionDeAfter.getPageSetup().getRightMargin(), 90.0); // 3.17 cm
        Assert.assertEquals(sectionDeAfter.getPageSetup().getTopMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionDeAfter.getPageSetup().getBottomMargin(), 72.0); // 2.54 cm
        Assert.assertEquals(sectionDeAfter.getPageSetup().getHeaderDistance(), 35.4); // 1.25 cm
        Assert.assertEquals(sectionDeAfter.getPageSetup().getFooterDistance(), 35.4); // 1.25 cm
        Assert.assertEquals(sectionDeAfter.getPageSetup().getTextColumns().getSpacing(), 35.4); // 1.25 cm
    }
}
