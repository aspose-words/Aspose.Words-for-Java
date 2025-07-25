// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.WordML2003SaveOptions;
import org.testng.Assert;
import com.aspose.ms.NUnit.Framework.msAssert;
import com.aspose.words.SaveFormat;
import com.aspose.ms.System.IO.File;
import com.aspose.ms.System.Environment;
import org.testng.annotations.DataProvider;


@Test
public class ExWordML2003SaveOptions extends ApiExampleBase
{
    @Test (dataProvider = "prettyFormatDataProvider")
    public void prettyFormat(boolean prettyFormat) throws Exception
    {
        //ExStart
        //ExFor:WordML2003SaveOptions
        //ExFor:WordML2003SaveOptions.SaveFormat
        //ExSummary:Shows how to manage output document's raw content.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
        // to modify how we save the document to the WordML save format.
        WordML2003SaveOptions options = new WordML2003SaveOptions();

        Assert.assertEquals(SaveFormat.WORD_ML, options.getSaveFormat());

        // Set the "PrettyFormat" property to "true" to apply tab character indentation and
        // newlines to make the output document's raw content easier to read.
        // Set the "PrettyFormat" property to "false" to save the document's raw content in one continuous body of the text.
        options.setPrettyFormat(prettyFormat);

        doc.save(getArtifactsDir() + "WordML2003SaveOptions.PrettyFormat.xml", options);

        String fileContents = File.readAllText(getArtifactsDir() + "WordML2003SaveOptions.PrettyFormat.xml");
        String newLine = Environment.getNewLine();
        if (prettyFormat)
            Assert.assertTrue(fileContents.contains(
                    $"<o:DocumentProperties>{newLine}\t\t" +
                        $"<o:Revision>1</o:Revision>{newLine}\t\t" +
                        $"<o:TotalTime>0</o:TotalTime>{newLine}\t\t" +
                        $"<o:Pages>1</o:Pages>{newLine}\t\t" +
                        $"<o:Words>0</o:Words>{newLine}\t\t" +
                        $"<o:Characters>0</o:Characters>{newLine}\t\t" +
                        $"<o:Lines>1</o:Lines>{newLine}\t\t" +
                        $"<o:Paragraphs>1</o:Paragraphs>{newLine}\t\t" +
                        $"<o:CharactersWithSpaces>0</o:CharactersWithSpaces>{newLine}\t\t" +
                        $"<o:Version>11.5606</o:Version>{newLine}\t" +
                    "</o:DocumentProperties>"));
        else
            Assert.assertTrue(fileContents.contains(
                    "<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>" +
                    "<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" +
                    "<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>"));
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "prettyFormatDataProvider")
	public static Object[][] prettyFormatDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}

    @Test (dataProvider = "memoryOptimizationDataProvider")
    public void memoryOptimization(boolean memoryOptimization) throws Exception
    {
        //ExStart
        //ExFor:WordML2003SaveOptions
        //ExSummary:Shows how to manage memory optimization.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello world!");

        // Create a "WordML2003SaveOptions" object to pass to the document's "Save" method
        // to modify how we save the document to the WordML save format.
        WordML2003SaveOptions options = new WordML2003SaveOptions();

        // Set the "MemoryOptimization" flag to "true" to decrease memory consumption
        // during the document's saving operation at the cost of a longer saving time.
        // Set the "MemoryOptimization" flag to "false" to save the document normally.
        options.setMemoryOptimization(memoryOptimization);

        doc.save(getArtifactsDir() + "WordML2003SaveOptions.MemoryOptimization.xml", options);
        //ExEnd
    }

	//JAVA-added data provider for test method
	@DataProvider(name = "memoryOptimizationDataProvider")
	public static Object[][] memoryOptimizationDataProvider() throws Exception
	{
		return new Object[][]
		{
			{false},
			{true},
		};
	}
}

