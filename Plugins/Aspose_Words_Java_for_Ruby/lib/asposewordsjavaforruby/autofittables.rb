module Asposewordsjavaforruby
  module AutoFitTables
    def initialize()
        # The path to the documents directory.
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Demonstrate autofitting a table to the window.
        autofit_table_to_window()

        # Demonstrate autofitting a table to its contents.
        autofit_table_to_contents()

        # Demonstrate autofitting a table to fixed column widths.
        autofit_table_to_fixed_column_widths()
    end

=begin
    ExStart
    ExFor:Table.AutoFit
    ExFor:AutoFitBehavior
    ExId:FitTableToPageWidth
    ExSummary:Autofits a table to fit the page width.
=end
    def autofit_table_to_window()
        #data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        # Open the document
        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + "TestFile.doc")
        
        node_type = Rjb::import('com.aspose.words.NodeType')
        table = doc.getChild(node_type.TABLE, 0, true)
        
        # Autofit the first table to the page width.
        autofit_behavior = Rjb::import("com.aspose.words.AutoFitBehavior")
        table.autoFit(autofit_behavior.AUTO_FIT_TO_WINDOW)

        # Save the document to disk.
        doc.save(@data_dir + "TestFile.AutoFitToWindow Out.doc")
        # ExEnd
        preferred_width_type = Rjb::import("com.aspose.words.PreferredWidthType")

        if (doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == preferred_width_type.PERCENT) then
            puts "PreferredWidth type is not percent."
        end

        if (doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getValue() == 100) then    
            puts "PreferredWidth value is different than 100."
        end
    end

=begin
    ExStart
    ExFor:Table.AutoFit
    ExFor:AutoFitBehavior
    ExId:FitTableToContents
    ExSummary:Autofits a table in the document to its contents.
=end
    def autofit_table_to_contents()
        # Open the document
        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + "TestFile.doc")
        
        node_type = Rjb::import('com.aspose.words.NodeType')
        table = doc.getChild(node_type.TABLE, 0, true)
        
        # Autofit the table to the cell contents
        autofit_behavior = Rjb::import("com.aspose.words.AutoFitBehavior")
        table.autoFit(autofit_behavior.AUTO_FIT_TO_CONTENTS)

        # Save the document to disk.
        doc.save(@data_dir + "TestFile.AutoFitToContents Out.doc")
        # ExEnd
        preferred_width_type = Rjb::import("com.aspose.words.PreferredWidthType")

        if (doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == preferred_width_type.AUTO) then
            puts "PreferredWidth type is not auto."
        end

        if (doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getType() == preferred_width_type.AUTO) then
            puts "PrefferedWidth on cell is not auto."
        end

        if(doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getPreferredWidth().getValue() == 0) then
            puts "PreferredWidth value is not 0."
        end
    end

=begin
    ExStart
    ExFor:Table.AutoFit
    ExFor:AutoFitBehavior
    ExId:DisableAutoFitAndUseFixedWidths
    ExSummary:Disables autofitting and enables fixed widths for the specified table.
=end
    def autofit_table_to_fixed_column_widths()
        # Open the document
        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + "TestFile.doc")
        
        node_type = Rjb::import('com.aspose.words.NodeType')
        table = doc.getChild(node_type.TABLE, 0, true)
        
        # Disable autofitting on this table.
        autofit_behavior = Rjb::import("com.aspose.words.AutoFitBehavior")
        table.autoFit(autofit_behavior.AUTO_FIT_TO_CONTENTS)

        # Save the document to disk.
        doc.save(@data_dir + "TestFile.FixedWidth Out.doc")
        # ExEnd
        preferred_width_type = Rjb::import("com.aspose.words.PreferredWidthType")

        if (doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getType() == preferred_width_type.AUTO) then
            puts "PreferredWidth type is not auto."
        end

        if (doc.getFirstSection().getBody().getTables().get(0).getPreferredWidth().getValue() == 0) then
            puts "PreferredWidth value is not 0."
        end

        if (doc.getFirstSection().getBody().getTables().get(0).getFirstRow().getFirstCell().getCellFormat().getWidth() == 0) then
            puts "Cell width is not correct."
        end
    end

  end
end
