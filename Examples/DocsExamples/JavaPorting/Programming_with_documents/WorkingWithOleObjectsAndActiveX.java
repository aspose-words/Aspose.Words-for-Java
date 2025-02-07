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
        //ExStart:InsertOleObject
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObjectInternal("http://www.aspose.com", "htmlfile", true, true, null);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObject.docx");
        //ExEnd:InsertOleObject
    }

    @Test
    public void insertOleObjectWithOlePackage() throws Exception
    {
        //ExStart:InsertOleObjectwithOlePackage
        //GistId:4996b573cf231d9f66ab0d1f3f981222
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

        //ExStart:GetAccessToOleObjectRawData
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Shape oleShape = (Shape) doc.getChild(NodeType.SHAPE, 0, true);
        byte[] oleRawData = oleShape.getOleFormat().getRawData();
        //ExEnd:GetAccessToOleObjectRawData
    }

    @Test
    public void insertOleObjectAsIcon() throws Exception
    {
        //ExStart:InsertOleObjectAsIcon
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObjectAsIcon(getMyDir() + "Presentation.pptx", false, getImagesDir() + "Logo icon.ico",
            "My embedded file");

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIcon.docx");
        //ExEnd:InsertOleObjectAsIcon
    }

    @Test
    public void insertOleObjectAsIconUsingStream() throws Exception
    {
        //ExStart:InsertOleObjectAsIconUsingStream
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        MemoryStream stream = new MemoryStream(File.readAllBytes(getMyDir() + "Presentation.pptx"));
        try /*JAVA: was using*/
    	{
            builder.insertOleObjectAsIconInternal(stream, "Package", getImagesDir() + "Logo icon.ico", "My embedded file");
    	}
        finally { if (stream != null) stream.close(); }

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOleObjectAsIconUsingStream.docx");
        //ExEnd:InsertOleObjectAsIconUsingStream
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

    @Test
    public void insertOnlineVideo() throws Exception
    {
        //ExStart:InsertOnlineVideo
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        String url = "https://youtu.be/t_1LYZ102RA";
        double width = 360.0;
        double height = 270.0;

        Shape shape = builder.insertOnlineVideo(url, width, height);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideo.docx");
        //ExEnd:InsertOnlineVideo
    }

    @Test
    public void insertOnlineVideoWithEmbedHtml() throws Exception
    {
        //ExStart:InsertOnlineVideoWithEmbedHtml
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        double width = 360.0;
        double height = 270.0;

        String videoUrl = "https://vimeo.com/52477838";
        String videoEmbedCode =
            "<iframe src=\"https://player.vimeo.com/video/52477838\" width=\"640\" height=\"360\" frameborder=\"0\" " +
            "title=\"Aspose\" webkitallowfullscreen mozallowfullscreen allowfullscreen></iframe>";

        byte[] thumbnailImageBytes = File.readAllBytes(getImagesDir() + "Logo.jpg");

        builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, width, height);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideoWithEmbedHtml.docx");
        //ExEnd:InsertOnlineVideoWithEmbedHtml
    }
}
