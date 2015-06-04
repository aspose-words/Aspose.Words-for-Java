package featurescomparison.workingwithtables.formattable.java;

import java.io.FileOutputStream;
import java.math.BigInteger;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHeight;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTString;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTVerticalJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STShd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STVerticalJc;

public class ApacheFormattedTable
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithtables/formattable/data/";

		// Create a new document from scratch
        XWPFDocument doc = new XWPFDocument();
    	// -- OR --
        // open an existing empty document with styles already defined
        //XWPFDocument doc = new XWPFDocument(new FileInputStream("base_document.docx"));

    	// Create a new table with 6 rows and 3 columns
    	int nRows = 6;
    	int nCols = 3;
        XWPFTable table = doc.createTable(nRows, nCols);

        // Set the table style. If the style is not defined, the table style
        // will become "Normal".
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        CTString styleStr = tblPr.addNewTblStyle();
        styleStr.setVal("StyledTable");

        // Get a list of the rows in the table
        List<XWPFTableRow> rows = table.getRows();
        int rowCt = 0;
        int colCt = 0;
        for (XWPFTableRow row : rows) {
        	// get table row properties (trPr)
        	CTTrPr trPr = row.getCtRow().addNewTrPr();
        	// set row height; units = twentieth of a point, 360 = 0.25"
        	CTHeight ht = trPr.addNewTrHeight();
        	ht.setVal(BigInteger.valueOf(360));

	        // get the cells in this row
        	List<XWPFTableCell> cells = row.getTableCells();
            // add content to each cell
        	for (XWPFTableCell cell : cells) {
        		// get a table cell properties element (tcPr)
        		CTTcPr tcpr = cell.getCTTc().addNewTcPr();
        		// set vertical alignment to "center"
        		CTVerticalJc va = tcpr.addNewVAlign();
        		va.setVal(STVerticalJc.CENTER);

        		// create cell color element
        		CTShd ctshd = tcpr.addNewShd();
                ctshd.setColor("auto");
                ctshd.setVal(STShd.CLEAR);
                if (rowCt == 0) {
                	// header row
                	ctshd.setFill("A7BFDE");
                }
            	else if (rowCt % 2 == 0) {
            		// even row
                	ctshd.setFill("D3DFEE");
            	}
            	else {
            		// odd row
                	ctshd.setFill("EDF2F8");
            	}

                // get 1st paragraph in cell's paragraph list
                XWPFParagraph para = cell.getParagraphs().get(0);
                // create a run to contain the content
                XWPFRun rh = para.createRun();
                // style cell as desired
                if (colCt == nCols - 1) {
                	// last column is 10pt Courier
                	rh.setFontSize(10);
                	rh.setFontFamily("Courier");
                }
                if (rowCt == 0) {
                	// header row
                    rh.setText("header row, col " + colCt);
                	rh.setBold(true);
                    para.setAlignment(ParagraphAlignment.CENTER);
                }
            	else if (rowCt % 2 == 0) {
            		// even row
                    rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
            	}
            	else {
            		// odd row
                    rh.setText("row " + rowCt + ", col " + colCt);
                    para.setAlignment(ParagraphAlignment.LEFT);
            	}
                colCt++;
        	} // for cell
        	colCt = 0;
        	rowCt++;
        } // for row

        // write the file
        FileOutputStream out = new FileOutputStream(dataPath + "Apache_styledTable_Out.docx");
        doc.write(out);
        out.close();
        
        System.out.println("Process Completed Successfully");
	}
}
