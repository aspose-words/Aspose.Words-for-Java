package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.Document;
import com.aspose.words.HorizontalAlignment;
import com.aspose.words.MeasurementUnits;
import com.aspose.words.examples.Utils;
import com.aspose.words.ShowInBalloons;

public class WorkingWithRevisionOptions {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertBetweenMeasurementUnits.class);

		SetMeasurementUnit(dataDir);
		SetRevisionBarsPosition(dataDir);

	}
	
	private static void SetMeasurementUnit(String dataDir) throws Exception {
		// ExStart:SetMeasurementUnit
		Document doc = new Document(dataDir + "Input.docx");

		// Set Measurement Units to Inches
		doc.getLayoutOptions().getRevisionOptions().setMeasurementUnit(MeasurementUnits.INCHES);
		// Show deletion revisions in balloon
		doc.getLayoutOptions().getRevisionOptions().setShowInBalloons(ShowInBalloons.FORMAT_AND_DELETE);
		// Show Comments
		doc.getLayoutOptions().setShowComments(true);

		doc.save(dataDir + "Revisions.SetMeasurementUnit_out.pdf");
		// ExEnd:SetMeasurementUnit
	}
	
	private static void SetRevisionBarsPosition(String dataDir) throws Exception {
		// ExStart:SetRevisionBarsPosition
		Document doc = new Document(dataDir + "Input.docx");

		//Renders revision bars on the right side of a page.
		doc.getLayoutOptions().getRevisionOptions().setRevisionBarsPosition(HorizontalAlignment.RIGHT);

		doc.save(dataDir + "Revisions.SetRevisionBarsPosition_out.pdf");
		// ExEnd:SetRevisionBarsPosition
	}

}
