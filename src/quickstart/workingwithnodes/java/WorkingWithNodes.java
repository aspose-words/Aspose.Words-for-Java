/*
 * Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */

package quickstart.workingwithnodes.java;

import com.aspose.words.*;

public class WorkingWithNodes
{
    public static void main(String[] args) throws Exception
    {
        // Create a new document.
        Document doc = new Document();

        // Creates and adds a paragraph node to the document.
        Paragraph para = new Paragraph(doc);

        // Typed access to the last section of the document.
        Section section = doc.getLastSection();
        section.getBody().appendChild(para);

        // Next print the node type of one of the nodes in the document.
        int nodeType = doc.getFirstSection().getBody().getNodeType();

        System.out.println("NodeType: " + Node.nodeTypeToString(nodeType));
    }
}




