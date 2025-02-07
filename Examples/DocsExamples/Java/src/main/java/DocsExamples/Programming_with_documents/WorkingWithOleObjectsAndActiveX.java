package DocsExamples.Programming_with_documents;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.io.FileUtils;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

@Test
public class WorkingWithOleObjectsAndActiveX extends DocsExamplesBase
{
    @Test
    public void insertOleObject() throws Exception
    {
        //ExStart:InsertOleObject
        //GistId:4996b573cf231d9f66ab0d1f3f981222
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.insertOleObject("http://www.aspose.com", "htmlfile", true, true, null);

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

        try(ByteArrayInputStream stream = new ByteArrayInputStream(FileUtils.readFileToByteArray(new File(getMyDir() + "Presentation.pptx"))))
    	{
            builder.insertOleObjectAsIcon(stream, "Package", getImagesDir() + "Logo icon.ico", "My embedded file");
    	}

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

        byte[] thumbnailImageBytes = Files.readAllBytes(Paths.get(getImagesDir() + "Logo.jpg"));

        builder.insertOnlineVideo(videoUrl, videoEmbedCode, thumbnailImageBytes, width, height);

        doc.save(getArtifactsDir() + "WorkingWithOleObjectsAndActiveX.InsertOnlineVideoWithEmbedHtml.docx");
        //ExEnd:InsertOnlineVideoWithEmbedHtml
    }
}
