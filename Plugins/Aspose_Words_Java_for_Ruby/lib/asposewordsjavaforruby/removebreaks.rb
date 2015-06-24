module Asposewordsjavaforruby
  module RemoveBreaks
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/document/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + "TestRemoveBreaks.doc")

        # Remove the page and section breaks from the document.
        # In Aspose.Words section breaks are represented as separate Section nodes in the document.
        # To remove these separate sections the sections are combined.
        remove_page_breaks(doc)
        #remove_section_breaks(doc)
        
        # Save the document.
        doc.save(data_dir + "TestRemoveBreaks Out.doc")
    end

    def remove_page_breaks(doc)
        # Retrieve all paragraphs in the document.
        node_type = Rjb::import("com.aspose.words.NodeType")
        paragraphs = doc.getChildNodes(node_type.PARAGRAPH, true)
        paragraphs_count = paragraphs.getCount()
       
        i = 0
        while (i < paragraphs_count) do
            paragraphs = doc.getChildNodes(node_type.PARAGRAPH, true)
            para = paragraphs.get(i)
            
            if (para.getParagraphFormat().getPageBreakBefore()) then
                para.getParagraphFormat().setPageBreakBefore(false)
            end

            runs = para.getRuns().toArray()
            runs.each do |run|
                control_char = Rjb::import("com.aspose.words.ControlChar")
                p run.getText().contains(control_char.PAGE_BREAK)
                abort()
                #if (run.getText().contains(control_char.PAGE_BREAK)) then
                    run_text = run.getText()
                    run_text = run_text.gsub(control_char.PAGE_BREAK, '')
                    run.setText(run_text)
                #end
            end
            i = i + 1
        end
    end

    def remove_section_breaks(doc)
        # Loop through all sections starting from the section that precedes the last one
        # and moving to the first section.
        i = doc.getSections().getCount()
        i = i - 2
        while (i >= 0) do
            # Copy the content of the current section to the beginning of the last section.
            doc.getLastSection().prependContent(doc.getSections().get(i))
            # Remove the copied section.
            doc.getSections().get(i).remove()
            i = i - 1
        end
    end    

  end
end
