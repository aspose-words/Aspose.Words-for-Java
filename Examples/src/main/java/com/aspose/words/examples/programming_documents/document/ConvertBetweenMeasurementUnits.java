
/*
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
package com.aspose.words.examples.programming_documents.document;

import com.aspose.words.ConvertUtil;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.PageSetup;
import com.aspose.words.examples.Utils;

public class ConvertBetweenMeasurementUnits {
	public static void main(String[] args) throws Exception {

		// ExStart:ConvertBetweenMeasurementUnits
		// The path to the documents directory.
		String dataDir = Utils.getDataDir(ConvertBetweenMeasurementUnits.class);

		// Open the document.
		Document doc = new Document();
		DocumentBuilder builder = new DocumentBuilder(doc);
		PageSetup pageSetup = builder.getPageSetup();
		pageSetup.setTopMargin(ConvertUtil.inchToPoint(1.0));
		pageSetup.setBottomMargin(ConvertUtil.inchToPoint(1.0));
		pageSetup.setLeftMargin(ConvertUtil.inchToPoint(1.5));
		pageSetup.setRightMargin(ConvertUtil.inchToPoint(1.5));
		pageSetup.setHeaderDistance(ConvertUtil.inchToPoint(0.2));
		pageSetup.setFooterDistance(ConvertUtil.inchToPoint(0.2));
		doc.save(dataDir + "output.doc");
		// ExEnd:ConvertBetweenMeasurementUnits

	}
}
