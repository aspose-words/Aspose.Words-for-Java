package com.aspose.words.examples.programming_documents.document;

import java.awt.Dimension;

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.GroupShape;
import com.aspose.words.Shape;
import com.aspose.words.ShapeType;
import com.aspose.words.examples.Utils;
import com.aspose.words.examples.programming_documents.tables.creation.BuildTableFromDataTable;

public class AddGroupShape {
	
	private static final String dataDir = Utils.getSharedDataDir(BuildTableFromDataTable.class) + "Document/";
	
	public static void main(String[] args) throws Exception {
		Document doc = new Document();
		doc.ensureMinimum();
		GroupShape gs = new GroupShape(doc);

		Shape shape = new Shape(doc, ShapeType.ACCENT_BORDER_CALLOUT_1);
		shape.setWidth(100);
		shape.setHeight(100);
		gs.appendChild(shape);

		Shape shape1 = new Shape(doc, ShapeType.ACTION_BUTTON_BEGINNING);
		shape1.setLeft(100);
		shape1.setWidth(100);
		shape1.setHeight(200);
		gs.appendChild(shape1);

		gs.setWidth(200);
		gs.setHeight(200);

		gs.setCoordSize(new Dimension(200, 200));

		DocumentBuilder builder = new DocumentBuilder(doc);
		builder.insertNode(gs);

		doc.save(dataDir + "AddGroupShape_out.docx");

	}

}
