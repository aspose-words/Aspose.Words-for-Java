package com.aspose.words.examples.document_object_model;

import com.aspose.words.Document;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.aspose.words.examples.Utils;

public class ChildNodes {

	public static void main(String[] args) throws Exception {
		String dataDir = Utils.getSharedDataDir(ChildNodes.class) + "DocumentObjectModel/";
		
		enumerateChildrenOfACompositeNodeUsingEnumeratorProvidedByChildNodesCollection(dataDir);
		enumerateChildrenOfACompositeNodeUsingIndexedAccess(dataDir);
	}
	
	public static void enumerateChildrenOfACompositeNodeUsingEnumeratorProvidedByChildNodesCollection(String dataDir) throws Exception {
		
		Document doc = new Document(dataDir + "Document.doc");          
		Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
		            
		NodeCollection children = paragraph.getChildNodes();
		for (Node child : (Iterable<Node>) children) {
		    // Paragraph may contain children of various types such as runs, shapes and so on.
		    if (child.getNodeType() == NodeType.RUN) {
		        // Say we found the node that we want, do something useful.
		        Run run = (Run)child;
		        System.out.println(run.getText());
		    }
		}
	}
	
	public static void enumerateChildrenOfACompositeNodeUsingIndexedAccess(String dataDir) throws Exception {
		
		Document doc = new Document(dataDir + "Document.doc");          
		Paragraph paragraph = (Paragraph)doc.getChild(NodeType.PARAGRAPH, 0, true);
		
		NodeCollection children = paragraph.getChildNodes();
		for (int i = 0; i < children.getCount(); i++) {
		    Node child = children.get(i);

		    // Paragraph may contain children of various types such as runs, shapes and so on.
		    if (child.getNodeType() == NodeType.RUN) {
		        // Say we found the node that we want, do something useful.
		        Run run = (Run)child;
		        System.out.println(run.getText());
		    }
		}
	}

}
