module Asposewordsjavaforruby
  module ExtractContentBasedOnStyles
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "Test.Styles.doc")

        # Define style names as they are specified in the Word document.
        para_style = "Heading 1"
        run_style = "Intense Emphasis"

        # Collect paragraphs with defined styles.
        # Show the number of collected paragraphs and display the text of this paragraphs.
        paragraphs = paragraphs_by_style_name(doc, para_style)
        para_size = paragraphs.size()
        #para_size = java_values($para_size)
        puts "Paragraphs with #{para_style} styles #{para_size}"
        
        paragraphs = paragraphs.toArray()
        save_format = Rjb::import("com.aspose.words.SaveFormat")
        
        paragraphs.each do |paragraph|
            puts paragraph.toString(save_format.TEXT)
        end    
        
        # Collect runs with defined styles.
        # Show the number of collected runs and display the text of this runs.
        runs = runs_by_style_name(doc, run_style)
        runs_size = runs.size()

        puts "Runs with #{run_style} styles #{runs_size}"
        runs = runs.toArray()
        runs.each do |run|
            puts run.getRange().getText()
        end   
    end

    def paragraphs_by_style_name(doc, para_style)
        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = Rjb::import("java.util.ArrayList").new
        # Get all paragraphs from the document.
        nodeType = Rjb::import("com.aspose.words.NodeType")
        paragraphs = doc.getChildNodes(nodeType.PARAGRAPH, true)
        paragraphs = paragraphs.toArray()
        
        # Look through all paragraphs to find those with the specified style.
        paragraphs.each do |paragraph|
            para_name = paragraph.getParagraphFormat().getStyle().getName()
            if (para_name == para_style) then
                paragraphsWithStyle.add(paragraph)
            end
        end
        paragraphsWithStyle
    end    

    def runs_by_style_name(doc, run_style)
        # Create an array to collect runs of the specified style.
        runsWithStyle = Rjb::import("java.util.ArrayList").new
        
        # Get all runs from the document.
        nodeType = Rjb::import("com.aspose.words.NodeType")
        runs = doc.getChildNodes(nodeType.RUN, true)
        
        # Look through all runs to find those with the specified style.
        runs = runs.toArray()
        runs.each do |run|
            if (run.getFont().getStyle().getName() == run_style) then
                runsWithStyle.add(run)
            end    
        end    
        runsWithStyle
    end    

  end
end
