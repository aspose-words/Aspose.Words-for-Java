package DocsExamples.Programming_with_Documents.Working_with_Document;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import java.util.Map;
import com.aspose.ms.System.msConsole;
import com.aspose.words.DocumentProperty;
import com.aspose.words.CustomDocumentProperties;
import com.aspose.ms.System.DateTime;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import com.aspose.words.ConvertUtil;
import com.aspose.words.ControlChar;


class DocumentPropertiesAndVariables extends DocsExamplesBase
{
    @Test
    public void getVariables() throws Exception
    {
        //ExStart:GetVariables
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document(getMyDir() + "Document.docx");
        
        String variables = "";
        for (Map.Entry<String, String> entry : doc.getVariables())
        {
            String name = entry.getKey();
            String value = entry.getValue();
            if ("".equals(variables))
                variables = "Name: " + name + "," + "Value: {1}" + value;
            else
                variables = variables + "Name: " + name + "," + "Value: {1}" + value;
        }

        System.out.println("\nDocument have following variables " + variables);
        //ExEnd:GetVariables
    }

    @Test
    public void enumerateProperties() throws Exception
    {
        //ExStart:EnumerateProperties
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document(getMyDir() + "Properties.docx");
        
        System.out.println("1. Document name: {0}",doc.getOriginalFileName());
        System.out.println("2. Built-in Properties");
        
        for (DocumentProperty prop : (Iterable<DocumentProperty>) doc.getBuiltInDocumentProperties())
            System.out.println("{0} : {1}",prop.getName(),prop.getValue());

        System.out.println("3. Custom Properties");
        
        for (DocumentProperty prop : (Iterable<DocumentProperty>) doc.getCustomDocumentProperties())
            System.out.println("{0} : {1}",prop.getName(),prop.getValue());
        //ExEnd:EnumerateProperties
    }

    @Test
    public void addCustomProperties() throws Exception
    {
        //ExStart:AddCustomProperties
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document(getMyDir() + "Properties.docx");

        CustomDocumentProperties customDocumentProperties = doc.getCustomDocumentProperties();
        
        if (customDocumentProperties.get("Authorized") != null) return;
        
        customDocumentProperties.add("Authorized", true);
        customDocumentProperties.add("Authorized By", "John Smith");
        customDocumentProperties.addInternal("Authorized Date", DateTime.getToday());
        customDocumentProperties.add("Authorized Revision", doc.getBuiltInDocumentProperties().getRevisionNumber());
        customDocumentProperties.add("Authorized Amount", 123.45);
        //ExEnd:AddCustomProperties
    }

    @Test
    public void removeCustomProperties() throws Exception
    {
        //ExStart:RemoveCustomProperties
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document(getMyDir() + "Properties.docx");
        doc.getCustomDocumentProperties().remove("Authorized Date");
        //ExEnd:RemoveCustomProperties
    }

    @Test
    public void removePersonalInformation() throws Exception
    {
        //ExStart:RemovePersonalInformation
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document(getMyDir() + "Properties.docx"); { doc.setRemovePersonalInformation(true); }

        doc.save(getArtifactsDir() + "DocumentPropertiesAndVariables.RemovePersonalInformation.docx");
        //ExEnd:RemovePersonalInformation
    }

    @Test
    public void configuringLinkToContent() throws Exception
    {
        //ExStart:ConfiguringLinkToContent
        //GistId:0593a8803015363f3026f648332e7026
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.startBookmark("MyBookmark");
        builder.writeln("Text inside a bookmark.");
        builder.endBookmark("MyBookmark");

        // Retrieve a list of all custom document properties from the file.
        CustomDocumentProperties customProperties = doc.getCustomDocumentProperties();
        // Add linked to content property.
        DocumentProperty customProperty = customProperties.addLinkToContent("Bookmark", "MyBookmark");
        customProperty = customProperties.get("Bookmark");

        boolean isLinkedToContent = customProperty.isLinkToContent();
        String linkSource = customProperty.getLinkSource();
        String customPropertyValue = customProperty.getValue().toString();
        //ExEnd:ConfiguringLinkToContent
    }

    @Test
    public void convertBetweenMeasurementUnits() throws Exception
    {
        //ExStart:ConvertBetweenMeasurementUnits
        //GistId:f266e937d2c656f9441071e9a7b053c1
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setTopMargin(ConvertUtil.inchToPoint(1.0));
        pageSetup.setBottomMargin(ConvertUtil.inchToPoint(1.0));
        pageSetup.setLeftMargin(ConvertUtil.inchToPoint(1.5));
        pageSetup.setRightMargin(ConvertUtil.inchToPoint(1.5));
        pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
        pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));
        //ExEnd:ConvertBetweenMeasurementUnits
    }

    @Test
    public void useControlCharacters()
    {
        //ExStart:UseControlCharacters
        //GistId:6269ddb6427f9ad20623d975774a615e
        final String TEXT = "test\r";
        // Replace "\r" control character with "\r\n".
        String replace = TEXT.replace(ControlChar.CR, ControlChar.CR_LF);
        //ExEnd:UseControlCharacters
    }
}
