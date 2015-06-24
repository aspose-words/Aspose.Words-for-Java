module Asposewordsjavaforruby
  module Bookmarks
    def initialize()
        # The path to the documents directory.
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/bookmarks/'
        
        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + 'TestDefect1352.doc')

        append_bookmark_text()

        # This perform the custom task of putting the row bookmark ends into the same row with the bookmark starts.
        untangle_row_bookmark(doc)
        
        # Now we can easily delete rows by a bookmark without damaging any other row's bookmarks.
        delete_row_by_bookmark(doc, 'ROW2')
    end
        
    def append_bookmark_text()
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/bookmarks/'
        
        # Load the source document.
        src_doc = Rjb::import('com.aspose.words.Document').new(data_dir + "Template.doc")

        # This is the bookmark whose content we want to copy.
        src_bookmark = src_doc.getRange().getBookmarks().get("ntf010145060")

        # We will be adding to this document.
        dst_doc = Rjb::import('com.aspose.words.Document').new()
        
        # Let's say we will be appending to the end of the body of the last section.
        #node_type = Rjb::import('com.aspose.words.NodeType')
        dst_node = dst_doc.getLastSection().getBody()

        # It is a good idea to use this import context object because multiple nodes are being imported.
        # If you import multiple times without a single context, it will result in many styles created.
        import_format_mode = Rjb::import('com.aspose.words.ImportFormatMode')
        importer = Rjb::import("com.aspose.words.NodeImporter").new(src_doc, dst_doc, import_format_mode.KEEP_SOURCE_FORMATTING)

        # This is the paragraph that contains the beginning of the bookmark.
        start_para = src_bookmark.getBookmarkStart().getParentNode()
        
        # This is the paragraph that contains the end of the bookmark.
        end_para = src_bookmark.getBookmarkEnd().getParentNode()

        if (start_para == "" || end_para == "") then
            raise "Parent of the bookmark start or end is not a paragraph, cannot handle this scenario yet."
        end

        util = Rjb::import("java.io.InputStream")
        # Limit ourselves to a reasonably simple scenario.
        spara = (start_para.getParentNode()).to_string
        epara = (end_para.getParentNode()).to_string
        #p spara.strip
        #abort('spara')
        
        
        if spara.strip != epara.strip then
        #if (start_para.getParentNode() != end_para.getParentNode()) then
            raise "Start and end paragraphs have different parents, cannot handle this scenario yet."
        end
        
        # We want to copy all paragraphs from the start paragraph up to (and including) the end paragraph,
        # therefore the node at which we stop is one after the end paragraph.
        endNode = end_para.getNextSibling()
        
        # This is the loop to go through all paragraph-level nodes in the bookmark.
        curNode = start_para
        cNode = curNode
        eNode = endNode
        
        while (cNode != eNode) do
            # This creates a copy of the current node and imports it (makes it valid) in the context
            # of the destination document. Importing means adjusting styles and list identifiers correctly.
            newNode = importer.importNode(curNode, true)
            curNode = curNode.getNextSibling()
            cNode = curNode
            dst_node.appendChild(newNode)
        end

        # Save the finished document.
        dst_doc.save(data_dir + "Template Out.doc");
    end

    def untangle_row_bookmark(doc)
        bookmarks = doc.getRange().getBookmarks()
        bookmarks_count = bookmarks.getCount()

        i = 0
        while i < bookmarks_count do
            bookmark = bookmarks.get(i)
            row1 = bookmark.getBookmarkStart().getAncestor(Rjb::import("com.aspose.words.Row"))
            row2 = bookmark.getBookmarkEnd().getAncestor(Rjb::import("com.aspose.words.Row"))

            # If both rows are found okay and the bookmark start and end are contained
            # in adjacent rows, then just move the bookmark end node to the end
            # of the last paragraph in the last cell of the top row.
            if ((row1 != "") && (row2 != "") && (row1.getNextSibling() == row2)) then
                row1.getLastCell().getLastParagraph().appendChild(bookmark.getBookmarkEnd())
            end
            $i +=1
        end

        # Save the document.
        doc.save(@data_dir + "TestDefect1352 Out.doc")
    end

    def delete_row_by_bookmark(doc, bookmark_name)
        raise 'bookmark_name not specified.' if bookmark_name.empty?

        bookmark = doc.getRange().getBookmarks().get(bookmark_name)
        
        if bookmark.nil? then
            return 
        end
        
        # Get the parent row of the bookmark. Exit if the bookmark is not in a row.
        row = bookmark.getBookmarkStart().getAncestor(Rjb::import('com.aspose.words.Row'))

        if row.nil? then
            return
        end
        
        # Remove the row.
        row.remove()

        # Save the document.
        doc.save(@data_dir + "TestDefect1352 Out.doc")
    end

  end
end
