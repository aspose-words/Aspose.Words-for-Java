package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;

@Test
public class WorkingWithOleObjectsAndActiveX extends DocsExamplesBase
{
    @Test
    public void insertOleObject() throws Exception
    {
        //ExStart:DocumentBuilderInsertOleObject
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
        //ExEnd:DocumentBuilderInsertOleObject
    }

    @Test
    public void insertOleObjectWithOlePackage() throws Exception
    {
        //ExStart:InsertOleObjectwithOlePackage
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] bs = FileUtils.readFileToByteArray(new File(getMyDir() + "Zip file.zip"));

        try (ByteArrayInputStream stream = new ByteArrayInputStream(bs))
        {
            Shape shape = builder.insertOleObject(stream, "Package", true, null);
            OlePackage olePackage = shape.getOleFormat().getOlePackage();
            olePackage.setFileName("filename.zip");
            olePackage.setDisplayName("displayname.zip");

            doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectWithOlePackage.docx");
        }
        //ExEnd:InsertOleObjectwithOlePackage

        //ExStart:GetAccessToOLEObjectRawData
        Shape oleShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        byte[] oleRawData = oleShape.getOleFormat().getRawData();
        //ExEnd:GetAccessToOLEObjectRawData
    }

    @Test
    public void insertOleObjectAsIcon() throws Exception
    {
        //ExStart:InsertOLEObjectAsIcon
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObjectAsIcon(getMyDir() + "Presentation.pptx", false, getImagesDir() + "Logo icon.ico",
            "My embedded file");

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
        //ExEnd:InsertOLEObjectAsIcon
    }

    @Test
    public void insertOleObjectAsIconUsingStream() throws Exception
    {
        //ExStart:InsertOLEObjectAsIconUsingStream
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        try(ByteArrayInputStream stream = new ByteArrayInputStream(FileUtils.readFileToByteArray(new File(getMyDir() + "Presentation.pptx"))))
    	{
            builder.insertOleObjectAsIcon(stream, "Package", getImagesDir() + "Logo icon.ico", "My embedded file");
    	}

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIconUsingStream.docx");
        //ExEnd:InsertOLEObjectAsIconUsingStream
    }

    @Test
    public void readActiveXControlProperties() throws Exception
    {
        Document doc = new Document(getMyDir() + "ActiveX controls.docx");

        String properties = "";
        for (Shape shape : (Iterable<Shape>) doc.getChildNodes(NodeType.SHAPE, true))
        {
            if (shape.getOleFormat() == null) break;

            OleControl oleControl = shape.getOleFormat().getOleControl();
            if (oleControl.isForms2OleControl())
            {
                Forms2OleControl checkBox = (Forms2OleControl) oleControl;
                properties = properties + "\nCaption: " + checkBox.getCaption();
                properties = properties + "\nValue: " + checkBox.getValue();
                properties = properties + "\nEnabled: " + checkBox.getEnabled();
                properties = properties + "\nType: " + checkBox.getType();
                if (checkBox.getChildNodes() != null)
                {
                    properties = properties + "\nChildNodes: " + checkBox.getChildNodes();
                }

                properties += "\n";
            }
        }

        properties = properties + "\nTotal ActiveX Controls found: " + doc.getChildNodes(NodeType.SHAPE, true).getCount();
        System.out.println("\n" + properties);
    }
}
