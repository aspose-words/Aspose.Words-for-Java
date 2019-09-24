package com.aspose.words.examples.programming_documents.tables.ColumnsAndRows;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;

public class KeepTablesAndRowsFromBreakingAcrossPages {

    private static final String dataDir = Utils.getSharedDataDir(KeepTablesAndRowsFromBreakingAcrossPages.class) + "Tables/";

    public static void main(String[] args) throws Exception {
        //ExStart:KeepTablesAndRowsFromBreakingAcrossPages
        // Keeping a Row from Breaking across Pages
        keepingARowFromBreakingAcrossPages();

        // Keeping a Table from Breaking across Pages
        keepingATableFromBreakingAcrossPages();
        //ExEnd:KeepTablesAndRowsFromBreakingAcrossPages
    }

    //ExStart:keepingARowFromBreakingAcrossPages
    public static void keepingARowFromBreakingAcrossPages() throws Exception {
        Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");

        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // Disable breaking across pages for all rows in the table.
        for (Row row : table) {
            row.getRowFormat().setAllowBreakAcrossPages(false);
        }

        doc.save(dataDir + "Table.DisableBreakAcrossPages_out.doc");
    }
    //ExEnd:keepingARowFromBreakingAcrossPages

    @SuppressWarnings("unchecked")
    //ExStart:keepingATableFromBreakingAcrossPages
    public static void keepingATableFromBreakingAcrossPages() throws Exception {
        Document doc = new Document(dataDir + "Table.TableAcrossPage.doc");
        // Retrieve the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        // To keep a table from breaking across a page we need to enable KeepWithNext
        // for every paragraph in the table except for the last paragraphs in the last
        // row of the table.
        for (Cell cell : (Iterable<Cell>) table.getChildNodes(NodeType.CELL, true)) {
            // Call this method if table's cell is created on the fly
            // newly created cell does not have paragraph inside
            cell.ensureMinimum();
            for (Paragraph para : cell.getParagraphs())
                if (!(cell.getParentRow().isLastRow() && para.isEndOfCell()))
                    para.getParagraphFormat().setKeepWithNext(true);
        }

        doc.save(dataDir + "Table.KeepTableTogether_out.doc");
    }
    //ExEnd:keepingATableFromBreakingAcrossPages
}
