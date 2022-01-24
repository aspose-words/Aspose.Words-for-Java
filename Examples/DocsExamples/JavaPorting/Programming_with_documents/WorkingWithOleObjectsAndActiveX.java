package DocsExamples.Programming_with_Documents;

// ********* THIS FILE IS AUTO PORTED *********

import com.aspose.ms.System.msString;
import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.IO.Stream;
import com.aspose.ms.System.IO.MemoryStream;
import com.aspose.words.Shape;
import com.aspose.words.OlePackage;
import com.aspose.words.NodeType;
import com.aspose.words.OleControl;
import com.aspose.words.Forms2OleControl;
import com.aspose.ms.System.msConsole;


class WorkingWithOleObjectsAndActiveX extends DocsExamplesBase
{
    @Test
    public void insertOleObject() throws Exception
    {
        //ExStart:DocumentBuilderInsertOleObject
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObjectInternal("http://www.aspose.com", "htmlfile", true, true, null);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
        //ExEnd:DocumentBuilderInsertOleObject
    }

    @Test
    public void insertOleObjectWithOlePackage() throws Exception
    {
        //ExStart:InsertOleObjectwithOlePackage
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        byte[] bs = File.readAllBytes(getMyDir() + "Zip file.zip");
        Stream stream = new MemoryStream(bs);
        try /*JAVA: was using*/
        {
            Shape shape = builder.insertOleObjectInternal(stream, "Package", true, null);
            OlePackage olePackage = shape.getOleFormat().getOlePackage();
            olePackage.setFileName("filename.zip");
            olePackage.setDisplayName("displayname.zip");

            doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectWithOlePackage.docx");
        }
        finally { if (stream != null) stream.close(); }
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

        MemoryStream stream = new MemoryStream(File.readAllBytes(getMyDir() + "Presentation.pptx"));
        try /*JAVA: was using*/
    	{
            builder.insertOleObjectAsIconInternal(stream, "Package", getImagesDir() + "Logo icon.ico", "My embedded file");
    	}
        finally { if (stream != null) stream.close(); }

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
            if (shape.getOleFormat() instanceof null) break;

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

                properties = msString.plusEqOperator(properties, "\n");
            }
        }

        properties = properties + "\nTotal ActiveX Controls found: " + doc.getChildNodes(NodeType.SHAPE, true).getCount();
        System.out.println("\n" + properties);
    }
}
