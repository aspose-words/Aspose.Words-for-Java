package com.aspose.words.examples.mail_merge;

import com.aspose.words.*;
import com.aspose.words.examples.Utils;
import com.aspose.words.net.System.Data.DataTable;

import javax.sql.rowset.CachedRowSet;
import javax.sql.rowset.RowSetMetaDataImpl;
import javax.sql.rowset.RowSetProvider;
import java.awt.*;
import java.sql.ResultSet;

//ExStart:
public class ApplyCustomFormattingDuringMailMerge {

    private static final String dataDir = Utils.getSharedDataDir(ApplyCustomFormattingDuringMailMerge.class) + "MailMerge/";

    public static void main(String[] args) throws Exception {
        Document doc = new Document(dataDir + "MailMerge.AlternatingRows.doc");

        // Add a handler for the MergeField event.
        doc.getMailMerge().setFieldMergingCallback(new HandleMergeFieldAlternatingRows());

        // Execute mail merge with regions.
        DataTable dataTable = getSuppliersDataTable();
        doc.getMailMerge().executeWithRegions(dataTable);

        doc.save(dataDir + "MailMerge.AlternatingRows Out.doc");
    }

    /**
     * Returns true if the value is odd; false if the value is even.
     */
    public static boolean isOdd(int value) throws Exception {
        return (value % 2 != 0);
    }

    /**
     * Create DataTable and fill it with data. In real life this DataTable
     * should be filled from a database.
     */
    private static DataTable getSuppliersDataTable() throws Exception {
        java.sql.ResultSet resultSet = createCachedRowSet(new String[]{"CompanyName", "ContactName"});

        for (int i = 0; i < 10; i++)
            addRow(resultSet, new String[]{"Company " + Integer.toString(i), "Contact " + Integer.toString(i)});

        return new DataTable(resultSet, "Suppliers");
    }

    /**
     * A helper method that creates an empty Java disconnected ResultSet with
     * the specified columns.
     */
    private static ResultSet createCachedRowSet(String[] columnNames) throws Exception {
        RowSetMetaDataImpl metaData = new RowSetMetaDataImpl();
        metaData.setColumnCount(columnNames.length);
        for (int i = 0; i < columnNames.length; i++) {
            metaData.setColumnName(i + 1, columnNames[i]);
            metaData.setColumnType(i + 1, java.sql.Types.VARCHAR);
        }

        CachedRowSet rowSet = RowSetProvider.newFactory().createCachedRowSet();
        ;
        rowSet.setMetaData(metaData);

        return rowSet;
    }

    /**
     * A helper method that adds a new row with the specified values to a
     * disconnected ResultSet.
     */
    private static void addRow(ResultSet resultSet, String[] values) throws Exception {
        resultSet.moveToInsertRow();

        for (int i = 0; i < values.length; i++)
            resultSet.updateString(i + 1, values[i]);

        resultSet.insertRow();

        // This "dance" is needed to add rows to the end of the result set properly.
        // If I do something else then rows are either added at the front or the result
        // set throws an exception about a deleted row during mail merge.
        resultSet.moveToCurrentRow();
        resultSet.last();
    }
}

class HandleMergeFieldAlternatingRows implements IFieldMergingCallback {
    /**
     * Called for every merge field encountered in the document. We can either
     * return some data to the mail merge engine or do something else with the
     * document. In this case we modify cell formatting.
     */
    public void fieldMerging(FieldMergingArgs e) throws Exception {
        if (mBuilder == null)
            mBuilder = new DocumentBuilder(e.getDocument());

        // This way we catch the beginning of a new row.
        if (e.getFieldName().equals("CompanyName")) {
            // Select the color depending on whether the row number is even or odd.
            Color rowColor;
            if (ApplyCustomFormattingDuringMailMerge.isOdd(mRowIdx))
                rowColor = new Color(213, 227, 235);
            else
                rowColor = new Color(242, 242, 242);

            // There is no way to set cell properties for the whole row at the moment,
            // so we have to iterate over all cells in the row.
            for (int colIdx = 0; colIdx < 4; colIdx++) {
                mBuilder.moveToCell(0, mRowIdx, colIdx, 0);
                mBuilder.getCellFormat().getShading().setBackgroundPatternColor(rowColor);
            }

            mRowIdx++;
        }
    }

    public void imageFieldMerging(ImageFieldMergingArgs args) throws Exception {
        // Do nothing.
    }

    private DocumentBuilder mBuilder;
    private int mRowIdx;
}
//ExEnd: