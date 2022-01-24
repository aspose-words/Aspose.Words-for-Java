package DocsExamples.Rendering_and_printing.Complex_examples_and_helpers;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.apache.commons.lang3.StringUtils;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.Stroke;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.text.MessageFormat;

@Test
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
        System.out.print(padding + layoutEnumerator.getType() + " - " + layoutEnumerator.getKind());

        if (layoutEnumerator.getType() == LayoutEntityType.SPAN)
            System.out.print(" - " + layoutEnumerator.getText());

        System.out.println();
    }

    /// <summary>
    /// Returns a string of spaces for padding purposes.
    /// </summary>
    private static String addPadding(String padding)
    {
        return padding + StringUtils.repeat(' ', 4);
    }
}

class OutlineLayoutEntitiesRenderer
{
    public static void run(Document doc, LayoutEnumerator layoutEnumerator, String folderPath) throws Exception {
        // Make sure the enumerator is at the beginning of the document.
        layoutEnumerator.reset();

        for (int pageIndex = 0; pageIndex < doc.getPageCount(); pageIndex++) {
            // Use the document class to find information about the current page.
            PageInfo pageInfo = doc.getPageInfo(pageIndex);

            final float RESOLUTION = 150.0f;
            Dimension pageSize = pageInfo.getSizeInPixels(1.0f, RESOLUTION);

            BufferedImage image = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_ARGB);

            Graphics2D graphics = image.createGraphics();
            try {
                // Make the background white.
                graphics.setBackground(Color.WHITE);
                graphics.clearRect(0, 0, image.getWidth(), image.getHeight());

                // Render the page to the graphics.
                doc.renderToScale(pageIndex, graphics, 0.0f, 0.0f, 1.0f);

                // Add an outline around each element on the page using the graphics object.
                addBoundingBoxToElementsOnPage(layoutEnumerator, graphics);

                // Move the enumerator to the next page if there is one.
                layoutEnumerator.moveNext();

                ImageIO.write(image, "png", new File(folderPath + MessageFormat.format("EnumerateLayoutElements.Page_{0}.png", pageIndex + 1)));
            } finally {
                if (graphics != null) graphics.dispose();
            }
        }
    }

    /// <summary>
    /// Adds a colored border around each layout element on the page.
    /// </summary>
    private static void addBoundingBoxToElementsOnPage(LayoutEnumerator layoutEnumerator, Graphics2D graphics) throws Exception
    {
        do
        {
            // Use MoveLastChild and MovePrevious to enumerate from last to the first enumeration is done backward,
            // so the lines of child entities are drawn first and don't overlap the parent's lines.
            if (layoutEnumerator.moveLastChild())
            {
                addBoundingBoxToElementsOnPage(layoutEnumerator, graphics);
                layoutEnumerator.moveParent();
            }

            Stroke stroke1 = new BasicStroke(1f);
            graphics.setColor(getColorFromType(layoutEnumerator.getType()));
            graphics.setStroke(stroke1);

            // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
            // Draw a line around the layout entity on the page.
            Rectangle2D.Float rectF = layoutEnumerator.getRectangle();
            graphics.drawRect((int) rectF.getX(), (int) rectF.getY(), (int) rectF.getWidth(), (int) rectF.getHeight());

            // Stop after all elements on the page have been processed.
            if (layoutEnumerator.getType() == LayoutEntityType.PAGE)
                return;
        } while (layoutEnumerator.movePrevious());
    }

    /// <summary>
    /// Returns a different colored pen for each entity type.
    /// </summary>
    private static Color getColorFromType(int type)
    {
        switch (type)
        {
            case LayoutEntityType.CELL:
                return Color.PINK;
            case LayoutEntityType.COLUMN:
                return Color.green;
            case LayoutEntityType.COMMENT:
                return Color.CYAN;
            case LayoutEntityType.ENDNOTE:
                return Color.lightGray;
            case LayoutEntityType.FOOTNOTE:
                return Color.lightGray;
            case LayoutEntityType.HEADER_FOOTER:
                return Color.DARK_GRAY;
            case LayoutEntityType.LINE:
                return Color.blue;
            case LayoutEntityType.NOTE_SEPARATOR:
                return Color.magenta;
            case LayoutEntityType.PAGE:
                return Color.RED;
            case LayoutEntityType.ROW:
                return Color.orange;
            case LayoutEntityType.SPAN:
                return Color.RED;
            case LayoutEntityType.TEXT_BOX:
                return Color.yellow;
            default:
                return Color.RED;
        }
    }

    /// <summary>
    /// Converts a value in points to pixels.
    /// </summary>
    private static int pointToPixel(float value, double resolution)
    {
        return (int) ConvertUtil.pointToPixel(value, resolution);
    }
}

