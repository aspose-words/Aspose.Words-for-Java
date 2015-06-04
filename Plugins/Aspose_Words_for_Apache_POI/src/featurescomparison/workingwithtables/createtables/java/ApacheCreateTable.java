package featurescomparison.workingwithtables.createtables.java;

import java.io.FileOutputStream;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class ApacheCreateTable
{
	public static void main(String[] args) throws Exception
	{
		String dataPath = "src/featurescomparison/workingwithtables/createtables/data/";
		
		XWPFDocument document = new XWPFDocument();
		
		// New 2x2 table
		XWPFTable tableOne = document.createTable();
		XWPFTableRow tableOneRowOne = tableOne.getRow(0);
		tableOneRowOne.getCell(0).setText("Hello");
		tableOneRowOne.addNewTableCell().setText("World");
		
		XWPFTableRow tableOneRowTwo = tableOne.createRow();
		tableOneRowTwo.getCell(0).setText("This is");
		tableOneRowTwo.getCell(1).setText("a table");
		
		// Add a break between the tables
		document.createParagraph().createRun().addBreak();
		
		// New 3x3 table
		XWPFTable tableTwo = document.createTable();
		XWPFTableRow tableTwoRowOne = tableTwo.getRow(0);
		tableTwoRowOne.getCell(0).setText("col one, row one");
		tableTwoRowOne.addNewTableCell().setText("col two, row one");
		tableTwoRowOne.addNewTableCell().setText("col three, row one");
		
		XWPFTableRow tableTwoRowTwo = tableTwo.createRow();
		tableTwoRowTwo.getCell(0).setText("col one, row two");
		tableTwoRowTwo.getCell(1).setText("col two, row two");
		tableTwoRowTwo.getCell(2).setText("col three, row two");
		
		XWPFTableRow tableTwoRowThree = tableTwo.createRow();
		tableTwoRowThree.getCell(0).setText("col one, row three");
		tableTwoRowThree.getCell(1).setText("col two, row three");
		tableTwoRowThree.getCell(2).setText("col three, row three");
		
		FileOutputStream outStream = new FileOutputStream(dataPath + "Apache_CreateTable_Out.doc");
		document.write(outStream);
		outStream.close();
		
		System.out.println("Process Completed Successfully");
	}
}
