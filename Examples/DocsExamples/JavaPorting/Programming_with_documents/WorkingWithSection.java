package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.Section;
import com.aspose.words.PaperSize;


class WorkingWithSection extends DocsExamplesBase
{
    @Test
    public void addSection() throws Exception
    {
        //ExStart:AddSection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        builder.writeln("Hello2");

        Section sectionToAdd = new Section(doc);
        doc.getSections().add(sectionToAdd);
        //ExEnd:AddSection
    }

    @Test
    public void deleteSection() throws Exception
    {
        //ExStart:DeleteSection
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello2");
        doc.appendChild(new Section(doc));

        doc.getSections().removeAt(0);
        //ExEnd:DeleteSection
    }

    @Test
    public void deleteAllSections() throws Exception
    {
        //ExStart:DeleteAllSections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello2");
        doc.appendChild(new Section(doc));

        doc.getSections().clear();
        //ExEnd:DeleteAllSections
    }

    @Test
    public void appendSectionContent() throws Exception
    {
        //ExStart:AppendSectionContent
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello22");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello3");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello45");

        // This is the section that we will append and prepend to.
        Section section = doc.getSections().get(2);

        // This copies the content of the 1st section and inserts it at the beginning of the specified section.
        Section sectionToPrepend = doc.getSections().get(0);
        section.prependContent(sectionToPrepend);

        // This copies the content of the 2nd section and inserts it at the end of the specified section.
        Section sectionToAppend = doc.getSections().get(1);
        section.appendContent(sectionToAppend);
        //ExEnd:AppendSectionContent
    }

    @Test
    public void cloneSection() throws Exception
    {
        //ExStart:CloneSection
        Document doc = new Document(getMyDir() + "Document.docx");
        Section cloneSection = doc.getSections().get(0).deepClone();
        //ExEnd:CloneSection
    }

    @Test
    public void copySection() throws Exception
    {
        //ExStart:CopySection
        Document srcDoc = new Document(getMyDir() + "Document.docx");
        Document dstDoc = new Document();

        Section sourceSection = srcDoc.getSections().get(0);
        Section newSection = (Section) dstDoc.importNode(sourceSection, true);
        dstDoc.getSections().add(newSection);
        
        dstDoc.save(getArtifactsDir() + "WorkingWithSection.CopySection.docx");
        //ExEnd:CopySection
    }

    @Test
    public void deleteHeaderFooterContent() throws Exception
    {
        //ExStart:DeleteHeaderFooterContent
        Document doc = new Document(getMyDir() + "Document.docx");
        
        Section section = doc.getSections().get(0);
        section.clearHeadersFooters();
        //ExEnd:DeleteHeaderFooterContent
    }

    @Test
    public void deleteSectionContent() throws Exception
    {
        //ExStart:DeleteSectionContent
        Document doc = new Document(getMyDir() + "Document.docx");
        
        Section section = doc.getSections().get(0);
        section.clearContent();
        //ExEnd:DeleteSectionContent
    }

    @Test
    public void modifyPageSetupInAllSections() throws Exception
    {
        //ExStart:ModifyPageSetupInAllSections
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.writeln("Hello1");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello22");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello3");
        doc.appendChild(new Section(doc));
        builder.writeln("Hello45");

        // It is important to understand that a document can contain many sections,
        // and each section has its page setup. In this case, we want to modify them all.
        for (Section section : (Iterable<Section>) doc)
            section.getPageSetup().setPaperSize(PaperSize.LETTER);

        doc.save(getArtifactsDir() + "WorkingWithSection.ModifyPageSetupInAllSections.doc");
        //ExEnd:ModifyPageSetupInAllSections
    }

    @Test
    public void sectionsAccessByIndex() throws Exception
    {
        //ExStart:SectionsAccessByIndex
        Document doc = new Document(getMyDir() + "Document.docx");
        
        Section section = doc.getSections().get(0);
        section.getPageSetup().setLeftMargin(90.0); // 3.17 cm
        section.getPageSetup().setRightMargin(90.0); // 3.17 cm
        section.getPageSetup().setTopMargin(72.0); // 2.54 cm
        section.getPageSetup().setBottomMargin(72.0); // 2.54 cm
        section.getPageSetup().setHeaderDistance(35.4); // 1.25 cm
        section.getPageSetup().setFooterDistance(35.4); // 1.25 cm
        section.getPageSetup().getTextColumns().setSpacing(35.4); // 1.25 cm
        //ExEnd:SectionsAccessByIndex
    }
}
