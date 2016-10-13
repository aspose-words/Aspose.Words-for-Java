package com.aspose.words.examples.programming_documents.fields;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import com.aspose.words.Document;
import com.aspose.words.FieldStart;
import com.aspose.words.FieldType;
import com.aspose.words.Node;
import com.aspose.words.NodeCollection;
import com.aspose.words.NodeType;
import com.aspose.words.Run;
import com.aspose.words.examples.Utils;

/**
 * Shows how to rename merge fields in a Word document.
 */
public class RenameMergeFields {

	private static final String dataDir = Utils.getSharedDataDir(RenameMergeFields.class) + "Fields/";

	/**
	 * Finds all merge fields in a Word document and changes their names.
	 */
	public static void main(String[] args) throws Exception {
		// Specify your document name here.
		Document doc = new Document(dataDir + "RenameMergeFields.doc");

		// Select all field start nodes so we can find the merge fields.
		NodeCollection fieldStarts = doc.getChildNodes(NodeType.FIELD_START, true);
		for (FieldStart fieldStart : (Iterable<FieldStart>) fieldStarts) {
			if (fieldStart.getFieldType() == FieldType.FIELD_MERGE_FIELD) {
				MergeField mergeField = new MergeField(fieldStart);
				mergeField.setName(mergeField.getName() + "_Renamed");
			}
		}
		doc.save(dataDir + "RenameMergeFields Out.doc");
	}
}

/**
 * Represents a facade object for a merge field in a Microsoft Word document.
 */
class MergeField {

	private final Node mFieldStart;
	private final Node mFieldSeparator;
	private final Node mFieldEnd;
	private static final Pattern G_REGEX = Pattern.compile("\\s*(MERGEFIELD\\s|)(\\s|)(\\S+)\\s+");

	MergeField(FieldStart fieldStart) throws Exception {
		if (fieldStart.equals(null))
			throw new IllegalArgumentException("fieldStart");
		if (fieldStart.getFieldType() != FieldType.FIELD_MERGE_FIELD)
			throw new IllegalArgumentException("Field start type must be FieldMergeField.");

		mFieldStart = fieldStart;

		// Find the field separator node.
		mFieldSeparator = fieldStart.getField().getSeparator();
		if (mFieldSeparator == null)
			throw new IllegalStateException("Cannot find field separator.");

		mFieldEnd = fieldStart.getField().getEnd();
	}

	/**
	 * Gets or sets the name of the merge field.
	 */
	String getName() throws Exception {
		return ((FieldStart) mFieldStart).getField().getResult().replace("«", "").replace("»", "");
	}

	void setName(String value) throws Exception {
		// Merge field name is stored in the field result which is a Run
		// node between field separator and field end.
		Run fieldResult = (Run) mFieldSeparator.getNextSibling();
		fieldResult.setText(java.text.MessageFormat.format("«{0}»", value));

		// But sometimes the field result can consist of more than one run, delete these runs.
		removeSameParent(fieldResult.getNextSibling(), mFieldEnd);

		updateFieldCode(value);
	}

	private void updateFieldCode(String fieldName) throws Exception {
		// Field code is stored in a Run node between field start and field separator.
		Run fieldCode = (Run) mFieldStart.getNextSibling();
		Matcher matcher = G_REGEX.matcher(((FieldStart) mFieldStart).getField().getFieldCode());

		matcher.find();

		String newFieldCode = java.text.MessageFormat.format(" {0}{1} ", matcher.group(1).toString(), fieldName);
		fieldCode.setText(newFieldCode);

		// But sometimes the field code can consist of more than one run, delete these runs.
		removeSameParent(fieldCode.getNextSibling(), mFieldSeparator);
	}

	/**
	 * Removes nodes from start up to but not including the end node. Start and
	 * end are assumed to have the same parent.
	 */
	private static void removeSameParent(Node startNode, Node endNode) throws Exception {
		if ((endNode != null) && (startNode.getParentNode() != endNode.getParentNode()))
			throw new IllegalArgumentException("Start and end nodes are expected to have the same parent.");

		Node curChild = startNode;
		while ((curChild != null) && (curChild != endNode)) {
			Node nextChild = curChild.getNextSibling();
			curChild.remove();
			curChild = nextChild;
		}
	}
}