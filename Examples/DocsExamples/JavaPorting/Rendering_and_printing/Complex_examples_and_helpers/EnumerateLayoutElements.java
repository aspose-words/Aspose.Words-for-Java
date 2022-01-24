package DocsExamples.Complex_examples_and_helpers;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.LayoutEnumerator;
import com.aspose.ms.System.msConsole;
import com.aspose.words.LayoutEntityType;
import com.aspose.ms.System.msString;
import com.aspose.words.PageInfo;
import java.awt.image.BufferedImage;
import com.aspose.ms.System.Drawing.msSize;
import java.awt.Graphics2D;
import java.awt.Color;
import com.aspose.ms.System.Drawing.RectangleF;
import com.aspose.ms.System.Drawing.Rectangle;
import com.aspose.ms.System.Convert;
import com.aspose.words.ConvertUtil;


public class EnumerateLayoutElements extends DocsExamplesBase
{
    @Test
    public void getLayoutElements() throws Exception
    {
        Document doc = new Document(getMyDir() + "Document layout.docx");

        // Enumerator which is used to "walk" the elements of a rendered document.
        LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

        // Use the enumerator to write information about each layout element to the console.
        LayoutInfoWriter.run(layoutEnumerator);

        // Adds a border around each layout element and saves each page as a JPEG image to the data directory.
        OutlineLayoutEntitiesRenderer.run(doc, layoutEnumerator, getArtifactsDir());
    }
}

class LayoutInfoWriter
{
    public static void run(LayoutEnumerator layoutEnumerator) throws Exception
    {
        displayLayoutElements(layoutEnumerator, "");
    }

    /// <summary>
    /// Enumerates forward through each layout element in the document and prints out details of each element. 
    /// </summary>
    private static void displayLayoutElements(LayoutEnumerator layoutEnumerator, String padding) throws Exception
    {
        do
        {
            displayEntityInfo(layoutEnumerator, padding);

            if (layoutEnumerator.moveFirstChild())
            {
                // Recurse into this child element.
                displayLayoutElements(layoutEnumerator, addPadding(padding));
                layoutEnumerator.moveParent();
            }
        } while (layoutEnumerator.moveNext());
    }

    /// <summary>
    /// Displays information about the current layout entity to the console.
    /// </summary>
    private static void displayEntityInfo(LayoutEnumerator layoutEnumerator, String padding) throws Exception
    {
        msConsole.write(padding + layoutEnumerator.getType() + " - " + layoutEnumerator.getKind());

        if (layoutEnumerator.getType() == LayoutEntityType.SPAN)
            msConsole.write(" - " + layoutEnumerator.getText());

        msConsole.writeLine();
    }

    /// <summary>
    /// Returns a string of spaces for padding purposes.
    /// </summary>
    private static String addPadding(String padding)
    {
        return padding + msString.newString(' ', 4);
    }
}

class OutlineLayoutEntitiesRenderer
{
    public static void run(Document doc, LayoutEnumerator layoutEnumerator, String folderPath) throws Exception
    {
        // Make sure the enumerator is at the beginning of the document.
        layoutEnumerator.reset();

        for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++)
        {
            // Use the document class to find information about the current page.
            PageInfo pageInfo = doc.getPageInfo(pageIndex);

            final float RESOLUTION = 150.0f;
            /*Size*/long pageSize = pageInfo.getSizeInPixelsInternal(1.0f, RESOLUTION);

            BufferedImage img = new BufferedImage(msSize.getWidth(pageSize), msSize.getHeight(pageSize));
            try /*JAVA: was using*/
            {
                img.SetResolution(RESOLUTION, RESOLUTION);

                Graphics2D g = Graphics2D.FromImage(img);
                try /*JAVA: was using*/
                {
                    // Make the background white.
                    g.Clear(Color.WHITE);

                    // Render the page to the graphics.
                    doc.renderToScaleInternal(pageIndex, g, 0.0f, 0.0f, 1.0f);

                    // Add an outline around each element on the page using the graphics object.
                    addBoundingBoxToElementsOnPage(layoutEnumerator, g);

                    // Move the enumerator to the next page if there is one.
                    layoutEnumerator.moveNext();

                    img.Save(folderPath + $"EnumerateLayoutElements.Page_{pageIndex + 1}.png");
                }
                finally { if (g != null) g.close(); }
            }
            finally { if (img != null) img.close(); }
        }
    }

    /// <summary>
    /// Adds a colored border around each layout element on the page.
    /// </summary>
    private static void addBoundingBoxToElementsOnPage(LayoutEnumerator layoutEnumerator, Graphics2D g) throws Exception
    {
        do
        {
            // Use MoveLastChild and MovePrevious to enumerate from last to the first enumeration is done backward,
            // so the lines of child entities are drawn first and don't overlap the parent's lines.
            if (layoutEnumerator.moveLastChild())
            {
                addBoundingBoxToElementsOnPage(layoutEnumerator, g);
                layoutEnumerator.moveParent();
            }

            // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
            RectangleF rectF = layoutEnumerator.getRectangleInternal();
            Rectangle rect = new Rectangle(pointToPixel(rectF.getLeft(), g.DpiX), pointToPixel(rectF.getTop(), g.DpiY),
                pointToPixel(rectF.getWidth(), g.DpiX), pointToPixel(rectF.getHeight(), g.DpiY));

            // Draw a line around the layout entity on the page.
            g.DrawRectangle(getColoredPenFromType(layoutEnumerator.getType()), rect);

            // Stop after all elements on the page have been processed.
            if (layoutEnumerator.getType() == LayoutEntityType.PAGE)
                return;
        } while (layoutEnumerator.movePrevious());
    }

    /// <summary>
    /// Returns a different colored pen for each entity type.
    /// </summary>
    private static Pen getColoredPenFromType(/*LayoutEntityType*/int type)
    {
        switch (type)
        {
            case LayoutEntityType.CELL:
                return Pens.Purple;
            case LayoutEntityType.COLUMN:
                return Pens.Green;
            case LayoutEntityType.COMMENT:
                return Pens.LightBlue;
            case LayoutEntityType.ENDNOTE:
                return Pens.DarkRed;
            case LayoutEntityType.FOOTNOTE:
                return Pens.DarkBlue;
            case LayoutEntityType.HEADER_FOOTER:
                return Pens.DarkGreen;
            case LayoutEntityType.LINE:
                return Pens.Blue;
            case LayoutEntityType.NOTE_SEPARATOR:
                return Pens.LightGreen;
            case LayoutEntityType.PAGE:
                return Pens.Red;
            case LayoutEntityType.ROW:
                return Pens.Orange;
            case LayoutEntityType.SPAN:
                return Pens.Red;
            case LayoutEntityType.TEXT_BOX:
                return Pens.Yellow;
            default:
                return Pens.Red;
        }
    }

    /// <summary>
    /// Converts a value in points to pixels.
    /// </summary>
    private static int pointToPixel(float value, double resolution)
    {
        return Convert.toInt32(ConvertUtil.pointToPixel(value, resolution));
    }
}

