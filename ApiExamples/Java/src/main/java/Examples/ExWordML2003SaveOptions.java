package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.SaveFormat;
import com.aspose.words.WordML2003SaveOptions;
import org.apache.commons.io.FileUtils;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.nio.charset.StandardCharsets;

@Test
public class ExWordML2003SaveOptions extends ApiExampleBase {
    @Test(dataProvider = "prettyFormatDataProvider")
    public void prettyFormat(boolean prettyFormat) throws Exception {
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

        String fileContents = FileUtils.readFileToString(new File(getArtifactsDir() + "WordML2003SaveOptions.PrettyFormat.xml"), StandardCharsets.UTF_8);

        if (prettyFormat)
            Assert.assertTrue(fileContents.contains(
                    "<o:DocumentProperties>\r\n\t\t" +
                            "<o:Revision>1</o:Revision>\r\n\t\t" +
                            "<o:TotalTime>0</o:TotalTime>\r\n\t\t" +
                            "<o:Pages>1</o:Pages>\r\n\t\t" +
                            "<o:Words>0</o:Words>\r\n\t\t" +
                            "<o:Characters>0</o:Characters>\r\n\t\t" +
                            "<o:Lines>1</o:Lines>\r\n\t\t" +
                            "<o:Paragraphs>1</o:Paragraphs>\r\n\t\t" +
                            "<o:CharactersWithSpaces>0</o:CharactersWithSpaces>\r\n\t\t" +
                            "<o:Version>11.5606</o:Version>\r\n\t" +
                            "</o:DocumentProperties>"));
        else
            Assert.assertTrue(fileContents.contains(
                    "<o:DocumentProperties><o:Revision>1</o:Revision><o:TotalTime>0</o:TotalTime><o:Pages>1</o:Pages>" +
                            "<o:Words>0</o:Words><o:Characters>0</o:Characters><o:Lines>1</o:Lines><o:Paragraphs>1</o:Paragraphs>" +
                            "<o:CharactersWithSpaces>0</o:CharactersWithSpaces><o:Version>11.5606</o:Version></o:DocumentProperties>"));
        //ExEnd
    }

    @DataProvider(name = "prettyFormatDataProvider")
    public static Object[][] prettyFormatDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }

    @Test(dataProvider = "memoryOptimizationDataProvider")
    public void memoryOptimization(boolean memoryOptimization) throws Exception {
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

    @DataProvider(name = "memoryOptimizationDataProvider")
    public static Object[][] memoryOptimizationDataProvider() {
        return new Object[][]
                {
                        {false},
                        {true},
                };
    }
}

