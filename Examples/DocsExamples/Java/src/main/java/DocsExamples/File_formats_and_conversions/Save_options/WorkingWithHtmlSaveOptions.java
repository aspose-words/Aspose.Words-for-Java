package DocsExamples.File_formats_and_conversions.Save_options;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Paths;

@Test
public class WorkingWithHtmlSaveOptions extends DocsExamplesBase
{
    @Test
    public void exportRoundtripInformation() throws Exception
    {
        //ExStart:ExportRoundtripInformation
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportRoundtripInformation(true); }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ExportRoundtripInformation.html", saveOptions);
        //ExEnd:ExportRoundtripInformation
    }

    @Test
    public void exportFontsAsBase64() throws Exception
    {
        //ExStart:ExportFontsAsBase64
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setExportFontsAsBase64(true); }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
        //ExEnd:ExportFontsAsBase64
    }

    @Test
    public void exportResources() throws Exception
    {
        //ExStart:ExportResources
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL);
            saveOptions.setExportFontResources(true);
            saveOptions.setResourceFolder(getArtifactsDir() + "Resources");
            saveOptions.setResourceFolderAlias("http://example.com/resources");
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ExportResources.html", saveOptions);
        //ExEnd:ExportResources
    }

    @Test
    public void convertMetafilesToPng() throws Exception
    {
        //ExStart:ConvertMetafilesToPng
        String html =
                "<html>\n                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>\n                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>\n                    </svg>\n                </html>";

        // Use 'ConvertSvgToEmf' to turn back the legacy behavior
        // where all SVG images loaded from an HTML document were converted to EMF.
        // Now SVG images are loaded without conversion
        // if the MS Word version specified in load options supports SVG images natively.
        HtmlLoadOptions loadOptions = new HtmlLoadOptions(); { loadOptions.setConvertSvgToEmf(true); }
        Charset charset = StandardCharsets.UTF_8;
        Document doc = new Document(new ByteArrayInputStream(html.getBytes(charset)), loadOptions);

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setMetafileFormat(HtmlMetafileFormat.PNG); }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ConvertMetafilesToPng.html", saveOptions);
        //ExEnd:ConvertMetafilesToPng
    }

    @Test
    public void convertMetafilesToSvg() throws Exception
    {
        //ExStart:ConvertMetafilesToSvg
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        
        builder.write("Here is an SVG image: ");
        builder.insertHtml(
            "<svg height='210' width='500'>\r\n                <polygon points='100,10 40,198 190,78 10,78 160,198' \r\n                    style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />\r\n            </svg> ");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(); { saveOptions.setMetafileFormat(HtmlMetafileFormat.SVG); }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ConvertMetafilesToSvg.html", saveOptions);
        //ExEnd:ConvertMetafilesToSvg
    }

    @Test
    public void addCssClassNamePrefix() throws Exception
    {
        //ExStart:AddCssClassNamePrefix
        Document doc = new Document(getMyDir() + "Rendering.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        {
            saveOptions.setCssStyleSheetType(CssStyleSheetType.EXTERNAL); saveOptions.setCssClassNamePrefix("pfx_");
        }
        
        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.AddCssClassNamePrefix.html", saveOptions);
        //ExEnd:AddCssClassNamePrefix
    }

    @Test
    public void exportCidUrlsForMhtmlResources() throws Exception
    {
        //ExStart:ExportCidUrlsForMhtmlResources
        Document doc = new Document(getMyDir() + "Content-ID.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.MHTML);
        {
            saveOptions.setPrettyFormat(true); saveOptions.setExportCidUrlsForMhtmlResources(true);
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ExportCidUrlsForMhtmlResources.mhtml", saveOptions);
        //ExEnd:ExportCidUrlsForMhtmlResources
    }

    @Test
    public void resolveFontNames() throws Exception
    {
        //ExStart:ResolveFontNames
        Document doc = new Document(getMyDir() + "Missing font.docx");

        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            saveOptions.setPrettyFormat(true); saveOptions.setResolveFontNames(true);
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ResolveFontNames.html", saveOptions);
        //ExEnd:ResolveFontNames
    }

    @Test
    public void exportTextInputFormFieldAsText() throws Exception
    {
        //ExStart:ExportTextInputFormFieldAsText
        Document doc = new Document(getMyDir() + "Rendering.docx");

        File imagesDir = new File(Paths.get(getArtifactsDir(), "Images").toString());

        // The folder specified needs to exist and should be empty.
        if (imagesDir.exists())
            imagesDir.delete();

        imagesDir.mkdir();

        // Set an option to export form fields as plain text, not as HTML input elements.
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.HTML);
        {
            saveOptions.setExportTextInputFormFieldAsText(true); saveOptions.setImagesFolder(imagesDir.getPath());
        }

        doc.save(getArtifactsDir() + "WorkingWithHtmlSaveOptions.ExportTextInputFormFieldAsText.html", saveOptions);
        //ExEnd:ExportTextInputFormFieldAsText
    }
}
