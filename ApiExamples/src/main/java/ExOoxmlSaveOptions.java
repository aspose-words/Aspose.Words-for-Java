//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.MsWordVersion;
import com.aspose.words.Shape;
import com.aspose.words.NodeType;
import org.testng.Assert;
import com.aspose.words.ShapeMarkupLanguage;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.SaveFormat;
import com.aspose.words.ListTemplate;
import com.aspose.words.List;
import com.aspose.words.BreakType;

import java.io.ByteArrayOutputStream;
import java.text.MessageFormat;
import java.util.Date;

public class ExOoxmlSaveOptions extends ApiExampleBase
{
    @Test
    public void iso29500Strict() throws Exception
    {
        //ExStart
        //ExFor:OoxmlCompliance
        //ExFor:OoxmlSaveOptions.Compliance
        //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        //Set Word2003 version for document, for inserting image as VML shape
        doc.getCompatibilityOptions().optimizeFor(MsWordVersion.WORD_2003);

        builder.insertImage(getImageDir() + "dotnet-logo.png");

        // Loop through all single shapes inside document.
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            System.out.println(shape.getMarkupLanguage());
            Assert.assertEquals(shape.getMarkupLanguage(), ShapeMarkupLanguage.VML);//ExSkip
        }

        //Iso29500_2008 does not allow VML shapes, so you need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); 
        saveOptions.setSaveFormat(SaveFormat.DOCX);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, saveOptions);

        //Assert that image have drawingML markup language
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            Assert.assertEquals(shape.getMarkupLanguage(), ShapeMarkupLanguage.DML);
        }
    }

    @Test
    public void restartingDocumentList() throws Exception
    {
        //ExStart
        //ExFor:List.IsRestartAtEachSection
        //ExSummary:Shows how to specify that the list has to be restarted at each section.
        Document doc = new Document();

        doc.getLists().add(ListTemplate.NUMBER_DEFAULT);

        List list = doc.getLists().get(0);

        // Set true to specify that the list has to be restarted at each section.
        list.isRestartAtEachSection(true);

        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.getListFormat().setList(list);

        for (int i = 1; i <= 45; i++)
        {
            builder.write(MessageFormat.format("List Item {0}\n",i));

            // Insert section break.
            if (i == 15 || i == 30)
                builder.insertBreak(BreakType.SECTION_BREAK_NEW_PAGE);
        }

        // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
        OoxmlSaveOptions options = new OoxmlSaveOptions();
        options.setCompliance(OoxmlCompliance.ISO_29500_2008_TRANSITIONAL);

        doc.save(getMyDir() + "\\Artifacts\\RestartingDocumentList.docx", options);
        //ExEnd
    }

    @Test
    public void updatingLastSavedTimeDocument() throws Exception
    {
        //ExStart
        //ExFor:SaveOptions.UpdateLastSavedTimeProperty
        //ExSummary:Shows how to update a document time property when you want to save it
        Document doc = new Document(getMyDir() + "Document.doc");

        //Get last saved time
        Date documentTimeBeforeSave = doc.getBuiltInDocumentProperties().getLastSavedTime();

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
        saveOptions.setUpdateLastSavedTimeProperty(true);
        //ExEnd

        ByteArrayOutputStream dstStream = new ByteArrayOutputStream();
        doc.save(dstStream, saveOptions);

        Date documentTimeAfterSave = doc.getBuiltInDocumentProperties().getLastSavedTime();

        Assert.assertFalse(documentTimeBeforeSave == documentTimeAfterSave);
    }
}
