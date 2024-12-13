package Examples;

//////////////////////////////////////////////////////////////////////////
// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

import com.aspose.words.*;
import com.aspose.words.net.System.Data.DataRow;
import com.aspose.words.net.System.Data.DataSet;
import com.aspose.words.net.System.Data.DataTable;
import jdk.nashorn.internal.runtime.regexp.joni.Regex;
import org.testng.Assert;
import org.testng.annotations.Test;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.Date;
import java.util.stream.Stream;

public class ExLowCode extends ApiExampleBase
{
    @Test
    public void mergeDocuments() throws Exception
    {
        //ExStart
        //ExFor:Merger.Merge(String, String[])
        //ExFor:Merger.Merge(String[], MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
        //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
        //ExFor:LowCode.MergeFormatMode
        //ExFor:LowCode.Merger
        //ExSummary:Shows how to merge documents into a single output document.
        //There is a several ways to merge documents:
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SimpleMerge.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" });

        OoxmlSaveOptions ooxmlSaveOptions = new OoxmlSaveOptions();
        ooxmlSaveOptions.setPassword("Aspose.Words");
        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveOptions.docx", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, ooxmlSaveOptions, MergeFormatMode.KEEP_SOURCE_FORMATTING);

        Merger.merge(getArtifactsDir() + "LowCode.MergeDocument.SaveFormat.pdf", new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, SaveFormat.PDF, MergeFormatMode.KEEP_SOURCE_LAYOUT);

        Document doc = Merger.merge(new String[] { getMyDir() + "Big document.docx", getMyDir() + "Tables.docx" }, MergeFormatMode.MERGE_FORMATTING);
        doc.save(getArtifactsDir() + "LowCode.MergeDocument.DocumentInstance.docx");
        //ExEnd
    }

    @Test
    public void mergeDocumentInstances() throws Exception
    {
        //ExStart:MergeDocumentInstances
        //GistId:f0964b777330b758f6b82330b040b24c
        //ExFor:Merger.Merge(Document[], MergeFormatMode)
        //ExSummary:Shows how to merge input documents to a single document instance.
        DocumentBuilder firstDoc = new DocumentBuilder();
        firstDoc.getFont().setSize(16.0);
        firstDoc.getFont().setColor(Color.BLUE);
        firstDoc.write("Hello first word!");

        DocumentBuilder secondDoc = new DocumentBuilder();
        secondDoc.write("Hello second word!");

        Document mergedDoc = Merger.merge(new Document[] { firstDoc.getDocument(), secondDoc.getDocument() }, MergeFormatMode.KEEP_SOURCE_LAYOUT);
        Assert.assertEquals("Hello first word!\fHello second word!\f", mergedDoc.getText());
        //ExEnd:MergeDocumentInstances
    }

    @Test
    public void convert() throws Exception
    {
        //ExStart:Convert
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.Convert(String, String)
        //ExFor:Converter.Convert(String, String, SaveFormat)
        //ExFor:Converter.Convert(String, String, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code.
        Converter.convert(getMyDir() + "Document.docx", getArtifactsDir() + "LowCode.Convert.pdf");

        Converter.convert(getMyDir() + "Document.docx", getArtifactsDir() + "LowCode.Convert.rtf", SaveFormat.RTF);

        OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
        Converter.convert(getMyDir() + "Document.doc", getArtifactsDir() + "LowCode.Convert.docx", saveOptions);
        //ExEnd:Convert
    }

    @Test
    public void convertStream() throws Exception
    {
        //ExStart:ConvertStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
        //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
        //ExSummary:Shows how to convert documents with a single line of code (Stream).
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.SaveFormat.docx"))
            {
                Converter.convert(streamIn, streamOut, SaveFormat.DOCX);
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(); { saveOptions.setPassword("Aspose.Words"); }
            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ConvertStream.SaveOptions.docx"))
            {
                Converter.convert(streamIn, streamOut1, saveOptions);
            }
        }
        //ExEnd:ConvertStream
    }

    @Test
    public void convertToImages() throws Exception
    {
        //ExStart:ConvertToImages
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(String, String)
        //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, String, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images.
        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.png");

        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.jpeg", SaveFormat.JPEG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        Converter.convertToImages(getMyDir() + "Big document.docx", getArtifactsDir() + "LowCode.ConvertToImages.png", imageSaveOptions);
        //ExEnd:ConvertToImages
    }

    @Test
    public void convertToImagesStream() throws Exception
    {
        //ExStart:ConvertToImagesStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(String, SaveFormat)
        //ExFor:Converter.ConvertToImages(String, ImageSaveOptions)
        //ExFor:Converter.ConvertToImages(Document, SaveFormat)
        //ExFor:Converter.ConvertToImages(Document, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images stream.
        InputStream[] streams = Converter.convertToImages(getMyDir() + "Big document.docx", SaveFormat.PNG);

        ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
        imageSaveOptions.setPageSet(new PageSet(1));
        streams = Converter.convertToImages(getMyDir() + "Big document.docx", imageSaveOptions);

        streams = Converter.convertToImages(new Document(getMyDir() + "Big document.docx"), SaveFormat.PNG);

        streams = Converter.convertToImages(new Document(getMyDir() + "Big document.docx"), imageSaveOptions);
        //ExEnd:ConvertToImagesStream
    }

    @Test
    public void convertToImagesFromStream() throws Exception
    {
        //ExStart:ConvertToImagesFromStream
        //GistId:0ede368e82d1e97d02e615a76923846b
        //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
        //ExFor:Converter.ConvertToImages(Stream, ImageSaveOptions)
        //ExSummary:Shows how to convert document to images from stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx"))
        {
            InputStream[] streams = Converter.convertToImages(streamIn, SaveFormat.JPEG);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.PNG);
            imageSaveOptions.setPageSet(new PageSet(1));
            streams = Converter.convertToImages(streamIn, imageSaveOptions);
        }
        //ExEnd:ConvertToImagesFromStream
    }

    @Test
    public void compareDocuments() throws Exception
    {
        //ExStart:CompareDocuments
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Comparer.Compare(String, String, String, String, DateTime)
        //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime)
        //ExFor:Comparer.Compare(String, String, String, String, DateTime, CompareOptions)
        //ExFor:Comparer.Compare(String, String, String, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to simple compare documents.
        // There is a several ways to compare documents:
        String firstDoc = getMyDir() + "Table column bookmarks.docx";
        String secondDoc = getMyDir() + "Table column bookmarks.doc";

        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.1.docx", "Author", new Date());
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.2.docx", SaveFormat.DOCX, "Author", new Date());
        CompareOptions options = new CompareOptions();
        options.setIgnoreCaseChanges(true);
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.3.docx", "Author", new Date(), options);
        Comparer.compare(firstDoc, secondDoc, getArtifactsDir() + "LowCode.CompareDocuments.4.docx", SaveFormat.DOCX, "Author", new Date(), options);
        //ExEnd:CompareDocuments
    }

    @Test
    public void compareStreamDocuments() throws Exception {
        //ExStart:CompareStreamDocuments
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime)
        //ExFor:Comparer.Compare(Stream, Stream, Stream, SaveFormat, String, DateTime, CompareOptions)
        //ExSummary:Shows how to compare documents from the stream.
        // There is a several ways to compare documents from the stream:
        try (FileInputStream firstStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.docx")) {
            try (FileInputStream secondStreamIn = new FileInputStream(getMyDir() + "Table column bookmarks.doc")) {
                try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.1.docx")) {
                    Comparer.compare(firstStreamIn, secondStreamIn, streamOut, SaveFormat.DOCX, "Author", new Date());
                }

                try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.CompareStreamDocuments.2.docx")) {
                    CompareOptions options = new CompareOptions();
                    options.setIgnoreCaseChanges(true);
                    Comparer.compare(firstStreamIn, secondStreamIn, streamOut1, SaveFormat.DOCX, "Author", new Date(), options);
                }
            }
        }
        //ExEnd:CompareStreamDocuments
    }

    @Test
    public void mailMerge() throws Exception
    {
        //ExStart:MailMerge
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(String, String, String[], Object[])
        //ExFor:MailMerger.Execute(String, String, SaveFormat, String[], Object[])
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, String[], Object[])
        //ExSummary:Shows how to do mail merge operation for a single record.
        // There is a several ways to do mail merge operation:
        String doc = getMyDir() + "Mail merge.doc";

        String[] fieldNames = new String[] { "FirstName", "Location", "SpecialCharsInName()" };
        String[] fieldValues = new String[] { "James Bond", "London", "Classified" };

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.1.docx", fieldNames, fieldValues);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.2.docx", SaveFormat.DOCX, fieldNames, fieldValues);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMerge.3.docx", SaveFormat.DOCX, options, fieldNames, fieldValues);
        //ExEnd:MailMerge
    }

    @Test
    public void mailMergeStream() throws Exception {
        //ExStart:MailMergeStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, String[], Object[])
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, String[], Object[])
        //ExSummary:Shows how to do mail merge operation for a single record from the stream.
        // There is a several ways to do mail merge operation using documents from the stream:
        String[] fieldNames = new String[]{"FirstName", "Location", "SpecialCharsInName()"};
        String[] fieldValues = new String[]{"James Bond", "London", "Classified"};

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStream.1.docx")) {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, fieldNames, fieldValues);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStream.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, options, fieldNames, fieldValues);
            }
        }
        //ExEnd:MailMergeStream
    }

    @Test
    public void mailMergeDataRow() throws Exception
    {
        //ExStart:MailMergeDataRow
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(String, String, DataRow)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataRow)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow.
        // There is a several ways to do mail merge operation from a DataRow:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[] { "James Bond", "London", "Classified" });
        DataRow dataRow = dataTable.getRows().get(0);

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.1.docx", dataRow);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.2.docx", SaveFormat.DOCX, dataRow);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataRow.3.docx", SaveFormat.DOCX, options, dataRow);
        //ExEnd:MailMergeDataRow
    }

    @Test
    public void mailMergeStreamDataRow() throws Exception {
        //ExStart:MailMergeStreamDataRow
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataRow)
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataRow)
        //ExSummary:Shows how to do mail merge operation from a DataRow using documents from the stream.
        // There is a several ways to do mail merge operation from a DataRow using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});
        DataRow dataRow = dataTable.getRows().get(0);

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.1.docx")) {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, dataRow);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamDataRow.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, options, dataRow);
            }
        }
        //ExEnd:MailMergeStreamDataRow
    }

    @Test
    public void mailMergeDataTable() throws Exception
    {
        //ExStart:MailMergeDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(String, String, DataTable)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, DataTable)
        //ExFor:MailMerger.Execute(String, String, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable.
        // There is a several ways to do mail merge operation from a DataTable:
        String doc = getMyDir() + "Mail merge.doc";

        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx", dataTable);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.execute(doc, getArtifactsDir() + "LowCode.MailMergeDataTable.3.docx", SaveFormat.DOCX, options, dataTable);
        //ExEnd:MailMergeDataTable
    }

    @Test
    public void mailMergeStreamDataTable() throws Exception {
        //ExStart:MailMergeStreamDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, DataTable)
        //ExFor:MailMerger.Execute(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable();
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("Location");
        dataTable.getColumns().add("SpecialCharsInName()");

        dataTable.getRows().add(new String[]{"James Bond", "London", "Classified"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeDataTable.1.docx")) {
                MailMerger.execute(streamIn, streamOut, SaveFormat.DOCX, dataTable);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.execute(streamIn, streamOut1, SaveFormat.DOCX, options, dataTable);
            }
        }
        //ExEnd:MailMergeStreamDataTable
    }

    @Test
    public void mailMergeWithRegionsDataTable() throws Exception
    {
        //ExStart:MailMergeWithRegionsDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable.
        // There is a several ways to do mail merge with regions operation from a DataTable:
        String doc = getMyDir() + "Mail merge with regions.docx";

        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[] { "John", "Doe" });
        dataTable.getRows().add(new Object[] { "", "" });
        dataTable.getRows().add(new Object[] { "Jane", "Doe" });

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.1.docx", dataTable);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.2.docx", SaveFormat.DOCX, dataTable);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataTable.3.docx", SaveFormat.DOCX, options, dataTable);
        //ExEnd:MailMergeWithRegionsDataTable
    }

    @Test
    public void mailMergeStreamWithRegionsDataTable() throws Exception {
        //ExStart:MailMergeStreamWithRegionsDataTable
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataTable)
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataTable)
        //ExSummary:Shows how to do mail merge with regions operation from a DataTable using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataTable using documents from the stream:
        DataTable dataTable = new DataTable("MyTable");
        dataTable.getColumns().add("FirstName");
        dataTable.getColumns().add("LastName");
        dataTable.getRows().add(new Object[]{"John", "Doe"});
        dataTable.getRows().add(new Object[]{"", ""});
        dataTable.getRows().add(new Object[]{"Jane", "Doe"});

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx")) {
                MailMerger.executeWithRegions(streamIn, streamOut, SaveFormat.DOCX, dataTable);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.executeWithRegions(streamIn, streamOut1, SaveFormat.DOCX, options, dataTable);
            }
        }
        //ExEnd:MailMergeStreamWithRegionsDataTable
    }

    @Test
    public void mailMergeWithRegionsDataSet() throws Exception
    {
        //ExStart:MailMergeWithRegionsDataSet
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(String, String, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(String, String, SaveFormat, MailMergeOptions, DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet.
        // There is a several ways to do mail merge with regions operation from a DataSet:
        String doc = getMyDir() + "Mail merge with regions data set.docx";

        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[] { 1, "John Doe" });
        tableCustomers.getRows().add(new Object[] { 2, "Jane Doe" });

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[] { 1, "Hawaiian", 2 });
        tableOrders.getRows().add(new Object[] { 2, "Pepperoni", 1 });
        tableOrders.getRows().add(new Object[] { 2, "Chicago", 1 });

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.1.docx", dataSet);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.2.docx", SaveFormat.DOCX, dataSet);
        MailMergeOptions options = new MailMergeOptions();
        options.setTrimWhitespaces(true);
        MailMerger.executeWithRegions(doc, getArtifactsDir() + "LowCode.MailMergeWithRegionsDataSet.3.docx", SaveFormat.DOCX, options, dataSet);
        //ExEnd:MailMergeWithRegionsDataSet
    }

    @Test
    public void mailMergeStreamWithRegionsDataSet() throws Exception {
        //ExStart:MailMergeStreamWithRegionsDataSet
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, DataSet)
        //ExFor:MailMerger.ExecuteWithRegions(Stream, Stream, SaveFormat, MailMergeOptions, DataSet)
        //ExSummary:Shows how to do mail merge with regions operation from a DataSet using documents from the stream.
        // There is a several ways to do mail merge with regions operation from a DataSet using documents from the stream:
        DataTable tableCustomers = new DataTable("Customers");
        tableCustomers.getColumns().add("CustomerID");
        tableCustomers.getColumns().add("CustomerName");
        tableCustomers.getRows().add(new Object[]{1, "John Doe"});
        tableCustomers.getRows().add(new Object[]{2, "Jane Doe"});

        DataTable tableOrders = new DataTable("Orders");
        tableOrders.getColumns().add("CustomerID");
        tableOrders.getColumns().add("ItemName");
        tableOrders.getColumns().add("Quantity");
        tableOrders.getRows().add(new Object[]{1, "Hawaiian", 2});
        tableOrders.getRows().add(new Object[]{2, "Pepperoni", 1});
        tableOrders.getRows().add(new Object[]{2, "Chicago", 1});

        DataSet dataSet = new DataSet();
        dataSet.getTables().add(tableCustomers);
        dataSet.getTables().add(tableOrders);
        dataSet.getRelations().add(tableCustomers.getColumns().get("CustomerID"), tableOrders.getColumns().get("CustomerID"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Mail merge.doc")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.1.docx")) {
                MailMerger.executeWithRegions(streamIn, streamOut, SaveFormat.DOCX, dataSet);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.MailMergeStreamWithRegionsDataTable.2.docx")) {
                MailMergeOptions options = new MailMergeOptions();
                options.setTrimWhitespaces(true);
                MailMerger.executeWithRegions(streamIn, streamOut1, SaveFormat.DOCX, options, dataSet);
            }
        }
        //ExEnd:MailMergeStreamWithRegionsDataSet
    }

    @Test
    public void replace() throws Exception
    {
        //ExStart:Replace
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(String, String, String, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, String, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document.
        // There is a several ways to replace string in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.1.docx", pattern, replacement);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.2.docx", SaveFormat.DOCX, pattern, replacement);
        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.Replace.3.docx", SaveFormat.DOCX, pattern, replacement, options);
        //ExEnd:Replace
    }

    @Test
    public void replaceStream() throws Exception {
        //ExStart:ReplaceStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String)
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, String, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string in the document using documents from the stream.
        // There is a several ways to replace string in the document using documents from the stream:
        String pattern = "(C)2006 Aspose Pty Ltd.";
        String replacement = "Copyright (C) 2024 by Aspose Pty Ltd.";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Footer.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStream.1.docx")) {
                Replacer.replace(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStream.2.docx")) {
                FindReplaceOptions options = new FindReplaceOptions();
                options.setFindWholeWordsOnly(false);
                Replacer.replace(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, options);
            }
        }
        //ExEnd:ReplaceStream
    }

    @Test
    public void replaceRegex() throws Exception
    {
        //ExStart:ReplaceRegex
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(String, String, Regex, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String)
        //ExFor:Replacer.Replace(String, String, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document.
        // There is a several ways to replace string with regex in the document:
        String doc = getMyDir() + "Footer.docx";
        String pattern = "gr(a|e)y";
        String replacement = "lavender";

        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.1.docx", pattern, replacement);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.2.docx", SaveFormat.DOCX, pattern, replacement);
        FindReplaceOptions options = new FindReplaceOptions();
        options.setFindWholeWordsOnly(false);
        Replacer.replace(doc, getArtifactsDir() + "LowCode.ReplaceRegex.3.docx", SaveFormat.DOCX, pattern, replacement, options);
        //ExEnd:ReplaceRegex
    }

    @Test
    public void replaceStreamRegex() throws Exception {
        //ExStart:ReplaceStreamRegex
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String)
        //ExFor:Replacer.Replace(Stream, Stream, SaveFormat, Regex, String, FindReplaceOptions)
        //ExSummary:Shows how to replace string with regex in the document using documents from the stream.
        // There is a several ways to replace string with regex in the document using documents from the stream:
        String pattern = "gr(a|e)y";
        String replacement = "lavender";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Replace regex.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.1.docx")) {
                Replacer.replace(streamIn, streamOut, SaveFormat.DOCX, pattern, replacement);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.ReplaceStreamRegex.2.docx")) {
                FindReplaceOptions options = new FindReplaceOptions();
                options.setFindWholeWordsOnly(false);
                Replacer.replace(streamIn, streamOut1, SaveFormat.DOCX, pattern, replacement, options);
            }
        }
        //ExEnd:ReplaceStreamRegex
    }

    //ExStart:BuildReportData
    //GistId:93fefe5344a8337b931d0fed5c028225
    //ExFor:ReportBuilder.BuildReport(String, String, Object)
    //ExFor:ReportBuilder.BuildReport(String, String, Object, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, ReportBuilderOptions)
    //ExSummary:Shows how to populate document with data.
    @Test //ExSkip
    public void buildReportData() throws Exception {
        // There is a several ways to populate document with data:
        String doc = getMyDir() + "Reporting engine template - If greedy (Java).docx";

        AsposeData obj = new AsposeData();
        {
            obj.setList(new ArrayList<>());
            {
                obj.getList().add("abc");
            }
        }

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.1.docx", obj);
        ReportBuilderOptions options = new ReportBuilderOptions();
        options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.2.docx", obj, options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.3.docx", SaveFormat.DOCX, obj);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportWithObject.4.docx", SaveFormat.DOCX, obj, options);
    }

    public static class AsposeData
    {
        public ArrayList<String> getList() { return mList; }; public void setList(ArrayList<String> value) { mList = value; };

        private ArrayList<String> mList;
    }
    //ExEnd:BuildReportData

    @Test
    public void buildReportDataStream() throws Exception {
        //ExStart:BuildReportDataStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data using documents from the stream.
        // There is a several ways to populate document with data using documents from the stream:
        AsposeData obj = new AsposeData();
        {
            obj.setList(new ArrayList<>());
            {
                obj.getList().add("abc");
            }
        }

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Reporting engine template - If greedy (Java).docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataStream.1.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut, SaveFormat.DOCX, obj);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataStream.2.docx")) {
                ReportBuilderOptions options = new ReportBuilderOptions();
                options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
                ReportBuilder.buildReport(streamIn, streamOut1, SaveFormat.DOCX, obj, options);
            }
        }
        //ExEnd:BuildReportDataStream
    }

    //ExStart:BuildReportDataSource
    //GistId:93fefe5344a8337b931d0fed5c028225
    //ExFor:ReportBuilder.BuildReport(String, String, Object, String)
    //ExFor:ReportBuilder.BuildReport(String, String, Object[], String[])
    //ExFor:ReportBuilder.BuildReport(String, String, Object, String, ReportBuilderOptions)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String)
    //ExFor:ReportBuilder.BuildReport(String, String, SaveFormat, Object, String, ReportBuilderOptions)
    //ExSummary:Shows how to populate document with data sources.
    @Test //ExSkip
    public void buildReportDataSource() throws Exception
    {
        // There is a several ways to populate document with data sources:
        String doc = getMyDir() + "Report building.docx";

        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.1.docx", sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.2.docx", new Object[] { sender }, new String[] { "s" });
        ReportBuilderOptions options = new ReportBuilderOptions();
        options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.3.docx", sender, "s", options);
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.4.docx", SaveFormat.DOCX, sender, "s");
        ReportBuilder.buildReport(doc, getArtifactsDir() + "LowCode.BuildReportDataSource.5.docx", SaveFormat.DOCX, sender, "s", options);
    }

    public static class MessageTestClass
    {
        public String getName() { return mName; }; public void setName(String value) { mName = value; };

        private String mName;
        public String getMessage() { return mMessage; }; public void setMessage(String value) { mMessage = value; };

        private String mMessage;

        public MessageTestClass(String name, String message)
        {
            setName(name);
            setMessage(message);
        }
    }
    //ExEnd:BuildReportDataSource

    @Test
    public void buildReportDataSourceStream() throws Exception {
        //ExStart:BuildReportDataSourceStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object[], String[])
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String)
        //ExFor:ReportBuilder.BuildReport(Stream, Stream, SaveFormat, Object, String, ReportBuilderOptions)
        //ExSummary:Shows how to populate document with data sources using documents from the stream.
        // There is a several ways to populate document with data sources using documents from the stream:
        MessageTestClass sender = new MessageTestClass("LINQ Reporting Engine", "Hello World");

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Report building.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.1.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut, SaveFormat.DOCX, new Object[]{sender}, new String[]{"s"});
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.2.docx")) {
                ReportBuilder.buildReport(streamIn, streamOut1, SaveFormat.DOCX, sender, "s");
            }

            try (FileOutputStream streamOut2 = new FileOutputStream(getArtifactsDir() + "LowCode.BuildReportDataSourceStream.3.docx")) {
                ReportBuilderOptions options = new ReportBuilderOptions();
                options.setOptions(ReportBuildOptions.ALLOW_MISSING_MEMBERS);
                ReportBuilder.buildReport(streamIn, streamOut2, SaveFormat.DOCX, sender, "s", options);
            }
        }
        //ExEnd:BuildReportDataSourceStream
    }

    @Test
    public void removeBlankPages() throws Exception
    {
        //ExStart:RemoveBlankPages
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.RemoveBlankPages(String, String)
        //ExFor:Splitter.RemoveBlankPages(String, String, SaveFormat)
        //ExSummary:Shows how to remove empty pages from the document.
        // There is a several ways to remove empty pages from the document:
        String doc = getMyDir() + "Blank pages.docx";

        Splitter.removeBlankPages(doc, getArtifactsDir() + "LowCode.RemoveBlankPages.1.docx");
        Splitter.removeBlankPages(doc, getArtifactsDir() + "LowCode.RemoveBlankPages.2.docx", SaveFormat.DOCX);
        //ExEnd:RemoveBlankPages
    }

    @Test
    public void removeBlankPagesStream() throws Exception {
        //ExStart:RemoveBlankPagesStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.RemoveBlankPages(Stream, Stream, SaveFormat)
        //ExSummary:Shows how to remove empty pages from the document from the stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Blank pages.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.RemoveBlankPagesStream.docx")) {
                Splitter.removeBlankPages(streamIn, streamOut, SaveFormat.DOCX);
            }
        }
        //ExEnd:RemoveBlankPagesStream
    }

    @Test
    public void extractPages() throws Exception
    {
        //ExStart:ExtractPages
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.ExtractPages(String, String, int, int)
        //ExFor:Splitter.ExtractPages(String, String, SaveFormat, int, int)
        //ExSummary:Shows how to extract pages from the document.
        // There is a several ways to extract pages from the document:
        String doc = getMyDir() + "Big document.docx";

        Splitter.extractPages(doc, getArtifactsDir() + "LowCode.ExtractPages.1.docx", 0, 2);
        Splitter.extractPages(doc, getArtifactsDir() + "LowCode.ExtractPages.2.docx", SaveFormat.DOCX, 0, 2);
        //ExEnd:ExtractPages
    }

    @Test
    public void extractPagesStream() throws Exception {
        //ExStart:ExtractPagesStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.ExtractPages(Stream, Stream, SaveFormat, int, int)
        //ExSummary:Shows how to extract pages from the document from the stream.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.ExtractPagesStream.docx")) {
                Splitter.extractPages(streamIn, streamOut, SaveFormat.DOCX, 0, 2);
            }
        }
        //ExEnd:ExtractPagesStream
    }

    @Test
    public void splitDocument() throws Exception
    {
        //ExStart:SplitDocument
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.Split(String, String, SplitOptions)
        //ExFor:Splitter.Split(String, String, SaveFormat, SplitOptions)
        //ExSummary:Shows how to split document by pages.
        String doc = getMyDir() + "Big document.docx";

        SplitOptions options = new SplitOptions();
        options.setSplitCriteria(SplitCriteria.PAGE);
        Splitter.split(doc, getArtifactsDir() + "LowCode.SplitDocument.1.docx", options);
        Splitter.split(doc, getArtifactsDir() + "LowCode.SplitDocument.2.docx", SaveFormat.DOCX, options);
        //ExEnd:SplitDocument
    }

    @Test
    public void splitDocumentStream() throws Exception {
        //ExStart:SplitDocumentStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Splitter.Split(Stream, SaveFormat, SplitOptions)
        //ExSummary:Shows how to split document from the stream by pages.
        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Big document.docx")) {
            SplitOptions options = new SplitOptions();
            options.setSplitCriteria(SplitCriteria.PAGE);
            InputStream[] stream = Splitter.split(streamIn, SaveFormat.DOCX, options);
        }
        //ExEnd:SplitDocumentStream
    }

    @Test
    public void watermarkText() throws Exception
    {
        //ExStart:WatermarkText
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetText(String, String, String)
        //ExFor:Watermarker.SetText(String, String, SaveFormat, String)
        //ExFor:Watermarker.SetText(String, String, String, TextWatermarkOptions)
        //ExFor:Watermarker.SetText(String, String, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document.
        String doc = getMyDir() + "Big document.docx";
        String watermarkText = "This is a watermark";

        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.1.docx", watermarkText);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.2.docx", SaveFormat.DOCX, watermarkText);
        TextWatermarkOptions options = new TextWatermarkOptions();
        options.setColor(Color.RED);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.3.docx", watermarkText, options);
        Watermarker.setText(doc, getArtifactsDir() + "LowCode.WatermarkText.4.docx", SaveFormat.DOCX, watermarkText, options);
        //ExEnd:WatermarkText
    }

    @Test
    public void watermarkTextStream() throws Exception {
        //ExStart:WatermarkTextStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String)
        //ExFor:Watermarker.SetText(Stream, Stream, SaveFormat, String, TextWatermarkOptions)
        //ExSummary:Shows how to insert watermark text to the document from the stream.
        String watermarkText = "This is a watermark";

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkTextStream.1.docx")) {
                Watermarker.setText(streamIn, streamOut, SaveFormat.DOCX, watermarkText);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.WatermarkTextStream.2.docx")) {
                TextWatermarkOptions options = new TextWatermarkOptions();
                options.setColor(Color.RED);
                Watermarker.setText(streamIn, streamOut1, SaveFormat.DOCX, watermarkText, options);
            }
        }
        //ExEnd:WatermarkTextStream
    }

    @Test
    public void watermarkImage() throws Exception
    {
        //ExStart:WatermarkImage
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetImage(String, String, String)
        //ExFor:Watermarker.SetImage(String, String, SaveFormat, String)
        //ExFor:Watermarker.SetImage(String, String, String, ImageWatermarkOptions)
        //ExFor:Watermarker.SetImage(String, String, SaveFormat, String, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document.
        String doc = getMyDir() + "Document.docx";
        String watermarkImage = getImageDir() + "Logo.jpg";

        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkImage.1.docx", watermarkImage);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.2.docx", SaveFormat.DOCX, watermarkImage);
        ImageWatermarkOptions options = new ImageWatermarkOptions();
        options.setScale(50.0);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.3.docx", watermarkImage, options);
        Watermarker.setImage(doc, getArtifactsDir() + "LowCode.SetWatermarkText.4.docx", SaveFormat.DOCX, watermarkImage, options);
        //ExEnd:WatermarkImage
    }

    @Test
    public void watermarkImageStream() throws Exception {
        //ExStart:WatermarkImageStream
        //GistId:93fefe5344a8337b931d0fed5c028225
        //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image)
        //ExFor:Watermarker.SetImage(Stream, Stream, SaveFormat, Image, ImageWatermarkOptions)
        //ExSummary:Shows how to insert watermark image to the document from a stream.
        BufferedImage image = ImageIO.read(new File(getImageDir() + "Logo.jpg"));

        try (FileInputStream streamIn = new FileInputStream(getMyDir() + "Document.docx")) {
            try (FileOutputStream streamOut = new FileOutputStream(getArtifactsDir() + "LowCode.SetWatermarkText.1.docx")) {
                Watermarker.setImage(streamIn, streamOut, SaveFormat.DOCX, image);
            }

            try (FileOutputStream streamOut1 = new FileOutputStream(getArtifactsDir() + "LowCode.SetWatermarkText.2.docx")) {
                ImageWatermarkOptions options = new ImageWatermarkOptions();
                options.setScale(50.0);
                Watermarker.setImage(streamIn, streamOut1, SaveFormat.DOCX, image, options);
            }
        }
        //ExEnd:WatermarkImageStream
    }
}

