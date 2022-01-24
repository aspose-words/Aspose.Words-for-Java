package DocsExamples.Programming_with_documents;

import com.aspose.words.Document;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.TextBox;
import org.testng.annotations.Test;

@Test
public class WorkingWithTextboxes
{
    @Test
    public void createALink() throws Exception
    {
        //ExStart:CreateALink
        Document doc = new Document();

        Shape shape1 = new Shape(doc, ShapeType.TEXT_BOX);
        Shape shape2 = new Shape(doc, ShapeType.TEXT_BOX);

        TextBox textBox1 = shape1.getTextBox();
        TextBox textBox2 = shape2.getTextBox();

        if (textBox1.isValidLinkTarget(textBox2))
            textBox1.setNext(textBox2);
        //ExEnd:CreateALink
    }

    @Test
    public void checkSequence() throws Exception
    {
        //ExStart:CheckSequence
        Document doc = new Document();

        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        TextBox textBox = shape.getTextBox();

        if (textBox.getNext() != null && textBox.getPrevious() == null)
        {
            System.out.println("The head of the sequence");
        }

        if (textBox.getNext() != null && textBox.getPrevious() != null)
        {
            System.out.println("The Middle of the sequence.");
        }

        if (textBox.getNext() == null && textBox.getPrevious() != null)
        {
            System.out.println("The Tail of the sequence.");
        }
        //ExEnd:CheckSequence
    }

    @Test
    public void breakALink() throws Exception
    {
        //ExStart:BreakALink
        Document doc = new Document();

        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        TextBox textBox = shape.getTextBox();

        // Break a forward link.
        textBox.breakForwardLink();

        // Break a forward link by setting a null.
        textBox.setNext(null);

        // Break a link, which leads to this textbox
        if (textBox.getPrevious() != null)
            textBox.getPrevious().breakForwardLink();
        //ExEnd:BreakALink
    }
}
