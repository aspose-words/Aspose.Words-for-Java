package DocsExamples.Programming_with_Documents.Working_with_Tables;

// ********* THIS FILE IS AUTO PORTED *********

import DocsExamples.DocsExamplesBase;
import org.testng.annotations.Test;
import com.aspose.words.Document;
import com.aspose.ms.System.msConsole;
import com.aspose.words.Table;
import com.aspose.words.NodeType;
import com.aspose.words.TableAlignment;
import com.aspose.words.BorderType;
import com.aspose.words.LineStyle;
import com.aspose.ms.System.Drawing.msColor;
import java.awt.Color;
import com.aspose.words.TextureIndex;
import com.aspose.words.Row;
import com.aspose.words.HeightRule;
import com.aspose.words.DocumentBuilder;
import com.aspose.words.RowFormat;
import com.aspose.words.Cell;
import com.aspose.words.TextOrientation;
import com.aspose.words.OoxmlSaveOptions;
import com.aspose.words.OoxmlCompliance;
import com.aspose.words.StyleIdentifier;
import com.aspose.words.TableStyleOptions;
import com.aspose.words.AutoFitBehavior;
import com.aspose.words.TableStyle;
import com.aspose.words.StyleType;
import com.aspose.words.CellFormat;


class WorkingWithTableStylesAndFormatting extends DocsExamplesBase
{
    @Test
    public void getDistanceBetweenTableSurroundingText() throws Exception
    {
        //ExStart:GetDistancebetweenTableSurroundingText
        Document doc = new Document(getMyDir() + "Tables.docx");

        System.out.println("\nGet distance between table left, right, bottom, top and the surrounding text.");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        msConsole.writeLine(table.getDistanceTop());
        msConsole.writeLine(table.getDistanceBottom());
        msConsole.writeLine(table.getDistanceRight());
        msConsole.writeLine(table.getDistanceLeft());
        //ExEnd:GetDistancebetweenTableSurroundingText
    }

    @Test
    public void applyOutlineBorder() throws Exception
    {
        //ExStart:ApplyOutlineBorder
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        // Align the table to the center of the page.
        table.setAlignment(TableAlignment.CENTER);
        // Clear any existing borders from the table.
        table.clearBorders();

        // Set a green border around the table but not inside.
        table.setBorder(BorderType.LEFT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.RIGHT, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.TOP, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);
        table.setBorder(BorderType.BOTTOM, LineStyle.SINGLE, 1.5, msColor.getGreen(), true);

        // Fill the cells with a light green solid color.
        table.setShading(TextureIndex.TEXTURE_SOLID, msColor.getLightGreen(), msColor.Empty);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.ApplyOutlineBorder.docx");
        //ExEnd:ApplyOutlineBorder
    }

    @Test
    public void buildTableWithBorders() throws Exception
    {
        //ExStart:BuildTableWithBorders
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        
        // Clear any existing borders from the table.
        table.clearBorders();
        
        // Set a green border around and inside the table.
        table.setBorders(LineStyle.SINGLE, 1.5, msColor.getGreen());

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.BuildTableWithBorders.docx");
        //ExEnd:BuildTableWithBorders
    }

    @Test
    public void modifyRowFormatting() throws Exception
    {
        //ExStart:ModifyRowFormatting
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
    public void setCellPadding() throws Exception
    {
        //ExStart:SetCellPadding
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.startTable();
        builder.insertCell();

        // Sets the amount of space (in points) to add to the left/top/right/bottom of the cell's contents.
        builder.getCellFormat().setPaddings(30.0, 50.0, 30.0, 50.0);
        builder.writeln("I'm a wonderful formatted cell.");

        builder.endRow();
        builder.endTable();

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.SetCellPadding.docx");
        //ExEnd:SetCellPadding
    }

    /// <summary>
    /// Shows how to modify formatting of a table cell.
    /// </summary>
    @Test
    public void modifyCellFormatting() throws Exception
    {
        //ExStart:ModifyCellFormatting
        Document doc = new Document(getMyDir() + "Tables.docx");
        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

        Cell firstCell = table.getFirstRow().getFirstCell();
        firstCell.getCellFormat().setWidth(30.0);
        firstCell.getCellFormat().setOrientation(TextOrientation.DOWNWARD);
        firstCell.getCellFormat().getShading().setForegroundPatternColor(msColor.getLightGreen());
        //ExEnd:ModifyCellFormatting
    }

    @Test
    public void formatTableAndCellWithDifferentBorders() throws Exception
    {
        //ExStart:FormatTableAndCellWithDifferentBorders
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
        builder.getCellFormat().getShading().setBackgroundPatternColor(msColor.getGreen());
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
    public void setTableTitleAndDescription() throws Exception
    {
        //ExStart:SetTableTitleAndDescription
        Document doc = new Document(getMyDir() + "Tables.docx");

        Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
        table.setTitle("Test title");
        table.setDescription("Test description");

        OoxmlSaveOptions options = new OoxmlSaveOptions(); { options.setCompliance(OoxmlCompliance.ISO_29500_2008_STRICT); }

        doc.getCompatibilityOptions().optimizeFor(com.aspose.words.MsWordVersion.WORD_2016);

        doc.save(getArtifactsDir() + "WorkingWithTableStylesAndFormatting.SetTableTitleAndDescription.docx", options);
        //ExEnd:SetTableTitleAndDescription
    }

    @Test
    public void allowCellSpacing() throws Exception
    {
        //ExStart:AllowCellSpacing
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
        tableStyle.getConditionalStyles().getFirstRow().getShading().setBackgroundPatternColor(msColor.getGreenYellow());
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
