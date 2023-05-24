package DocsExamples.Programming_with_documents.Working_with_tables;

import DocsExamples.DocsExamplesBase;
import com.aspose.words.*;
import org.testng.annotations.Test;

import java.awt.*;

@Test
public class WorkingWithTableStylesAndFormatting extends DocsExamplesBase
{
    @Test
    public void distanceBetweenTableSurroundingText() throws Exception
    {
        //ExStart:DistanceBetweenTableSurroundingText
        //GistId:8df1ad0825619cab7c80b571c6e6ba99
        Document doc = new Document(getMyDir() + "Tables.docx");

        System.out.println("\nGet distance between table left, right, bottom, top and the surrounding text.");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        System.out.println(table.getDistanceTop());
        System.out.println(table.getDistanceBottom());
        System.out.println(table.getDistanceRight());
        System.out.println(table.getDistanceLeft());
        //ExEnd:DistanceBetweenTableSurroundingText
    }

    @Test
    public void applyOutlineBorder() throws Exception
    {
        //ExStart:ApplyOutlineBorder
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        //ExStart:InlineTablePosition
        //GistId:8df1ad0825619cab7c80b571c6e6ba99
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        // Align the table to the center of the page.
        table.setAlignment(TableAlignment.CENTER);
        //ExEnd:InlineTablePosition
        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around the table but not inside.
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, Color.GREEN, true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, Color.GREEN, true);

        // Fill the cells with a light green solid color.
        table.setShading(TextureIndex.TEXTURE_SOLID, Color.lightGray, new Color(0, true));

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.ApplyOutlineBorder.docx");
        //ExEnd:ApplyOutlineBorder
    }

    @Test
    public void buildTableWithBorders() throws Exception
    {
        //ExStart:BuildTableWithBorders
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        
        // Clear any existing borders from the table.
        table.clearBorders();
        
        // Set a green border around and inside the table.
        table.setBorders(LineStyle.SINGLE, 1.5, Color.GREEN);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.BuildTableWithBorders.docx");
        //ExEnd:BuildTableWithBorders
    }

    @Test
    public void modifyRowFormatting() throws Exception
    {
        //ExStart:ModifyRowFormatting
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        
        // Retrieve the first row in the table.
        Row firstRow = table.getFirstRow();
        firstRow.getRowFormat().getBorders().setLineStyle(LineStyle.NONE);
        firstRow.getRowFormat().setHeightRule(HeightRule.AUTO);
        firstRow.getRowFormat().setAllowBreakAcrossPages(true);
        //ExEnd:ModifyRowFormatting
    }

    @Test
    public void applyRowFormatting() throws Exception
    {
        //ExStart:ApplyRowFormatting
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100.0);
        rowFormat.setHeightRule(HeightRule.EXACTLY);
        
        // These formatting properties are set on the table and are applied to all rows in the table.
        table.setLeftPadding(30.0);
        table.setRightPadding(30.0);
        table.setTopPadding(30.0);
        table.setBottomPadding(30.0);

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.ApplyRowFormatting.docx");
        //ExEnd:ApplyRowFormatting
    }

    @Test
    public void cellPadding() throws Exception
    {
        //ExStart:CellPadding
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        // Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
        builder.getCellFormat().setPaddings(30.0, 50.0, 30.0, 50.0);
        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.CellPadding.docx");
        //ExEnd:CellPadding
    }

    /// <summary>
    /// Shows how to modify formatting of a table cell.
    /// </summary>
    @Test
    public void modifyCellFormatting() throws Exception
    {
        //ExStart:ModifyCellFormatting
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Cell firstCell = table.getFirstRow().getFirstCell();
        firstCell.getCellFormat().setWidth(30.0);
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(Color.GREEN);
        //ExEnd:ModifyCellFormatting
    }

    @Test
    public void formatTableAndCellWithDifferentBorders() throws Exception
    {
        //ExStart:FormatTableAndCellWithDifferentBorders
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        // Set the borders for the entire table.
        table.setBorders(LineStyle.SINGLE, 2.0, Color.BLACK);
        
        // Set the cell shading for this cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.RED);
        builder.writeln("Cell #1");

        builder.insertCell();
        
        // Specify a different cell shading for the second cell.
        builder.getCellFormat().getShading().setBackgroundPatternColor(Color.GREEN);
        builder.writeln("Cell #2");

        builder.endRow();

        // Clear the cell formatting from previous operations.
        builder.getCellFormat().clearFormatting();

        builder.insertCell();

        // Create larger borders for the first cell of this row. This will be different
        // compared to the borders set for the table.
        builder.getCellFormat().getBorders().getLeft().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getRight().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getTop().setLineWidth(4.0);
        builder.getCellFormat().getBorders().getBottom().setLineWidth(4.0);
        builder.writeln("Cell #3");

        builder.insertCell();
        builder.getCellFormat().clearFormatting();
        builder.writeln("Cell #4");
        
        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.FormatTableAndCellWithDifferentBorders.docx");
        //ExEnd:FormatTableAndCellWithDifferentBorders
    }

    @Test
    public void tableTitleAndDescription() throws Exception
    {
        //ExStart:TableTitleAndDescription
        //GistId:458eb4fd5bd1de8b06fab4d1ef1acdc6
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setTitle("Test title");
        table.setDescription("Test description");

        OoxmlSaveOptions options = new OoxmlSaveOptions(); { options.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.getCompatibilityOptions().optimizeFor(com.aspose.words.MsWordVersion.WORD_2016);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.SetTableTitleAndDescription.docx", options);
        //ExEnd:TableTitleAndDescription
    }

    @Test
    public void allowCellSpacing() throws Exception
    {
        //ExStart:AllowCellSpacing
        //GistId:770bf20bd617f3cb80031a74cc6c9b73
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setAllowCellSpacing(true);
        table.setCellSpacing(2.0);
        
        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.AllowCellSpacing.docx");
        //ExEnd:AllowCellSpacing
    }

    @Test
    public void buildTableWithStyle() throws Exception
    {
        //ExStart:BuildTableWithStyle
        //GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        
        // We must insert at least one row first before setting any table formatting.
        builder.insertCell();

        // Set the table style used based on the unique style identifier.
        table.setStyleIdentifier(StyleIdentifier.MEDIUM_SHADING_1_ACCENT_1);
        
        // Apply which features should be formatted by the style.
        table.setStyleOptions(TableStyleOptions.FIRST_COLUMN | TableStyleOptions.ROW_BANDS | TableStyleOptions.FIRST_ROW);
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_CONTENTS);

        builder.writeln("Item");
        builder.getCellFormat().setRightPadding(40.0);
        builder.insertCell();
        builder.writeln("Quantity (kg)");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Apples");
        builder.insertCell();
        builder.writeln("20");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Bananas");
        builder.insertCell();
        builder.writeln("40");
        builder.endRow();

        builder.insertCell();
        builder.writeln("Carrots");
        builder.insertCell();
        builder.writeln("50");
        builder.endRow();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.BuildTableWithStyle.docx");
        //ExEnd:BuildTableWithStyle
    }

    @Test
    public void expandFormattingOnCellsAndRowFromStyle() throws Exception
    {
        //ExStart:ExpandFormattingOnCellsAndRowFromStyle
        //GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        Document doc = new Document(getMyDir() + "Tables.docx");

        // Get the first cell of the first table in the document.
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        Cell firstCell = table.getFirstRow().getFirstCell();

        // First print the color of the cell shading.
        // This should be empty as the current shading is stored in the table style.
        Color cellShadingBefore = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading before style expansion: " + cellShadingBefore);

        doc.expandTableStylesToDirectFormatting();

        // Now print the cell shading after expanding table styles.
        // A blue background pattern color should have been applied from the table style.
        Color cellShadingAfter = firstCell.getCellFormat().getShading().getBackgroundPatternColor();
        System.out.println("Cell shading after style expansion: " + cellShadingAfter);
        //ExEnd:ExpandFormattingOnCellsAndRowFromStyle
    }

    @Test
    public void createTableStyle() throws Exception
    {
        //ExStart:CreateTableStyle
        //GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Name");
        builder.insertCell();
        builder.write("Value");
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();

        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getBorders().setLineStyle(LineStyle.DOUBLE);
        tableStyle.getBorders().setLineWidth(1.0);
        tableStyle.setLeftPadding(18.0);
        tableStyle.setRightPadding(18.0);
        tableStyle.setTopPadding(12.0);
        tableStyle.setBottomPadding(12.0);

        table.setStyle(tableStyle);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.CreateTableStyle.docx");
        //ExEnd:CreateTableStyle
    }

    @Test
    public void defineConditionalFormatting() throws Exception
    {
        //ExStart:DefineConditionalFormatting
        //GistId:93b92a7e6f2f4bbfd9177dd7fcecbd8c
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();
        builder.write("Name");
        builder.insertCell();
        builder.write("Value");
        builder.endRow();
        builder.insertCell();
        builder.insertCell();
        builder.endTable();

        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.TABLE, "MyTableStyle1");
        tableStyle.getConditionalStyles().getFirstRow().getShading().setBackgroundPatternColor(Color.yellow);
        tableStyle.getConditionalStyles().getFirstRow().getShading().setTexture(TextureIndex.TEXTURE_NONE);

        table.setStyle(tableStyle);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.DefineConditionalFormatting.docx");
        //ExEnd:DefineConditionalFormatting
    }

    @Test
    public void setTableCellFormatting() throws Exception
    {
        //ExStart:DocumentBuilderSetTableCellFormatting
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        CellFormat cellFormat = builder.getCellFormat();
        cellFormat.setWidth(250.0);
        cellFormat.setLeftPadding(30.0);
        cellFormat.setRightPadding(30.0);
        cellFormat.setTopPadding(30.0);
        cellFormat.setBottomPadding(30.0);

        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.DocumentBuilderSetTableCellFormatting.docx");
        //ExEnd:DocumentBuilderSetTableCellFormatting
    }

    @Test
    public void setTableRowFormatting() throws Exception
    {
        //ExStart:DocumentBuilderSetTableRowFormatting
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        Table table = builder.startTable();
        builder.insertCell();

        RowFormat rowFormat = builder.getRowFormat();
        rowFormat.setHeight(100.0);
        rowFormat.setHeightRule(HeightRule.EXACTLY);
        
        // These formatting properties are set on the table and are applied to all rows in the table.
        table.setLeftPadding(30.0);
        table.setRightPadding(30.0);
        table.setTopPadding(30.0);
        table.setBottomPadding(30.0);

        builder.writeln("I'm a wonderful formatted row.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.DocumentBuilderSetTableRowFormatting.docx");
        //ExEnd:DocumentBuilderSetTableRowFormatting
    }
}
