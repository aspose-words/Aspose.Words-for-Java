//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.OdtSaveOptions;
import com.aspose.words.OdtSaveMeasureUnit;

public class ExOdtSaveOptions extends ApiExampleBase
{
    @Test
    public void measureUnitOption() throws Exception
    {
        //ExStart
        //ExFor:OdtSaveOptions.MeasureUnit
        //ExSummary: Show how to work with units of measure of document content
        Document doc = new Document(getMyDir() + "OdtSaveOptions.MeasureUnit.docx");

        //Open Office uses centimeters, MS Office uses inches
        OdtSaveOptions saveOptions = new OdtSaveOptions();
        saveOptions.setMeasureUnit(OdtSaveMeasureUnit.INCHES);

        doc.save(getArtifactsDir() + "OdtSaveOptions.MeasureUnit.odt");
        //ExEnd
    }
}
