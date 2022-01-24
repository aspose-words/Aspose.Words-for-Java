package DocsExamples.Programming_with_Documents.Contents_Management;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.StyleCollection;
import com.aspose.words.Style;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Theme;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.StyleType;
import com.aspose.words.StyleIdentifier;


class WorkingWithStylesAndThemes extends DocsExamplesBase
{
    @Test
    public void accessStyles() throws Exception
    {
        //ExStart:AccessStyles
        Document doc = new Document();

        String styleName = "";

        // Get styles collection from the document.
        StyleCollection styles = doc.getStyles();
        for (Style style : styles)
        {
            if ("".equals(styleName))
            {
                styleName = style.getName();
                System.out.println(styleName);
            }
            else
            {
                styleName = styleName + ", " + style.getName();
                System.out.println(styleName);
            }
        }
        //ExEnd:AccessStyles
    }

    @Test
    public void copyStyles() throws Exception
    {
        //ExStart:CopyStyles
        Document doc = new Document();
        Document target = new Document(getMyDir() + "Rendering.docx");

        target.copyStylesFromTemplate(doc);

        doc.save(getArtifactsDir() + "WorkingWithStylesAndThemes.CopyStyles.docx");
        //ExEnd:CopyStyles
    }

    @Test
    public void getThemeProperties() throws Exception
    {
        //ExStart:GetThemeProperties
        Document doc = new Document();

        Theme theme = doc.getTheme();

        System.out.println(theme.getMajorFonts().getLatin());
        System.out.println(theme.getMinorFonts().getEastAsian());
        System.out.println(theme.getColors().getAccent1());
        //ExEnd:GetThemeProperties 
    }

    @Test
    public void setThemeProperties() throws Exception
    {
        //ExStart:SetThemeProperties
        Document doc = new Document();

        Theme theme = doc.getTheme();
        theme.getMinorFonts().setLatin("Times New Roman");
        theme.getColors().setHyperlink(msColor.getGold());
        //ExEnd:SetThemeProperties 
    }

    @Test
    public void insertStyleSeparator() throws Exception
    {
        //ExStart:InsertStyleSeparator
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Style paraStyle = builder.getDocument().getStyles().add(StyleType.PARAGRAPH, "MyParaStyle");
        paraStyle.getFont().setBold(false);
        paraStyle.getFont().setSize(8.0);
        paraStyle.getFont().setName("Arial");

        // Append text with "Heading 1" style.
        builder.getParagraphFormat().setStyleIdentifier(StyleIdentifier.HEADING_1);
        builder.write("Heading 1");
        builder.insertStyleSeparator();

        // Append text with another style.
        builder.getParagraphFormat().setStyleName(paraStyle.getName());
        builder.write("This is text with some other formatting ");

        doc.save(getArtifactsDir() + "WorkingWithStylesAndThemes.InsertStyleSeparator.docx");
        //ExEnd:InsertStyleSeparator
    }
}
