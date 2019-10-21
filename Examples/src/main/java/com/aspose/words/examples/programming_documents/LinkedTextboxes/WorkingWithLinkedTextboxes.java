package com.aspose.words.examples.programming_documents.LinkedTextboxes;

import com.aspose.words.Document;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.TextBox;
import com.aspose.words.examples.Utils;

public class WorkingWithLinkedTextboxes {

    public static void main(String[] args) throws Exception {
        // TODO Auto-generated method stub
        String dataDir = Utils.getDataDir(WorkingWithLinkedTextboxes.class);

        CreateALink(dataDir);
        CheckSequence(dataDir);
        BreakALink(dataDir);
    }

    private static void CreateALink(String dataDir) throws Exception {
        // ExStart:CreateALink
        Document doc = new Document();
        Shape shape1 = new Shape(doc, ShapeType.TEXT_BOX);
        Shape shape2 = new Shape(doc, ShapeType.TEXT_BOX);

        TextBox textBox1 = shape1.getTextBox();
        TextBox textBox2 = shape2.getTextBox();

        if (textBox1.isValidLinkTarget(textBox2))
            textBox1.setNext(textBox2);
        // ExEnd:CreateALink
    }

    private static void CheckSequence(String dataDir) throws Exception {
        // ExStart:CheckSequence
        Document doc = new Document();
        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        TextBox textBox = shape.getTextBox();

        if ((textBox.getNext() != null) && (textBox.getPrevious() == null)) {
            System.out.println("The head of the sequence");
        }

        if ((textBox.getNext() != null) && (textBox.getPrevious() != null)) {
            System.out.println("The Middle of the sequence.");
        }

        if ((textBox.getNext() == null) && (textBox.getPrevious() != null)) {
            System.out.println("The Tail of the sequence.");
        }
        // ExEnd:CheckSequence
    }

    private static void BreakALink(String dataDir) throws Exception {
        // ExStart:BreakALink
        Document doc = new Document();
        Shape shape = new Shape(doc, ShapeType.TEXT_BOX);
        TextBox textBox = shape.getTextBox();

        // Break a forward link
        textBox.breakForwardLink();

        // Break a forward link by setting a null
        textBox.setNext(null);

        // Break a link, which leads to this textbox
        if (textBox.getPrevious() != null)
            textBox.getPrevious().breakForwardLink();
        // ExEnd:BreakALink
    }

}
