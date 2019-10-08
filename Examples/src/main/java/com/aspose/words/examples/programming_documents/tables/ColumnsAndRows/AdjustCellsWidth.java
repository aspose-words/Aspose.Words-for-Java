package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.Cell;
import com.aspose.words.CellMerge;
import com.aspose.words.Document;
import com.aspose.words.DocumentBase;
import com.aspose.words.Node;
import com.aspose.words.NodeType;
import com.aspose.words.Paragraph;
import com.aspose.words.Row;
import com.aspose.words.Run;
import com.aspose.words.Table;
import com.aspose.words.examples.Utils;

public class AdjustCellsWidth {

	//ExStart: AdjustCellsWidth
	private static java.util.List<Cell> _startMergeCells = new java.util.ArrayList<Cell>(3);
	private static final String dataDir = Utils.getSharedDataDir(WorkingWithColumns.class) + "Tables/";
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		Document doc = new Document(dataDir + "InputDoc.docx");
		Table table = (Table)doc.getChild(NodeType.TABLE, 0, true);

		Row row = new Row(doc);
		table.getRows().add(row);
		for (int i = 0; i < 3; i++)
		    appendOneCellAndAddValueToRow(doc, row, i + "—FirstRow", CellMerge.FIRST);


		Row row2 = new Row(doc);
		table.getRows().add(row2);
		for (int i = 0; i < 3; i++)
		    appendOneCellAndAddValueToRow(doc, row2, i + "—Row", CellMerge.PREVIOUS);

		doc.save(dataDir + "out.docx");
	}

	public static void appendOneCellAndAddValueToRow(Document doc, Row row, String value,int cellMerge) {
	    Cell cell = new Cell(doc);
	    cell.ensureMinimum();
	    cell.getFirstParagraph().getRuns().clear();
	    cell.getCellFormat().setVerticalMerge(cellMerge);
	    row.appendChild(cell);

	    if (cellMerge == CellMerge.FIRST)
	        _startMergeCells.add(cell);

	    InsertContent(cell, value);

	}
	
	private static void InsertContent(Cell cell, String value)
	{
	    Cell effectiveCell = cell;
	    DocumentBase doc = cell.getDocument();
	    Paragraph para = cell.getFirstParagraph();

	    if (cell.getCellFormat().getVerticalMerge() == CellMerge.PREVIOUS)
	    {
	        para = new Paragraph(doc);
	        int cellIndex = GetCellIndex(cell);

	        // The consolidated area is used to display the contents of the first vertically merged cell.
	        // So, move content to the first cell of merged range.
	        effectiveCell = _startMergeCells.get(cellIndex);
	    }

	    para.getRuns().add(new Run(doc, value));
	    if (para.getParentNode() == null)
	        effectiveCell.appendChild(para);
	}
	
	private static int GetCellIndex(Cell cell)
	{
	    Node nextCell = cell.getParentRow().getFirstCell();
	    int index = -1;

	    while (null != nextCell)
	    {
	        if (nextCell.getNodeType() == NodeType.CELL)
	            ++index;

	        if (cell == nextCell)
	            break;

	        nextCell = nextCell.getNextSibling();
	    }
	    return index;
	}
	//ExEnd: AdjustCellsWidth
}
