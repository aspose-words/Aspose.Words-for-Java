module Asposewordsjavaforruby
  module ExtractContent
    def initialize()
        # The path to the documents directory.
        @data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/document/'

        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(@data_dir + "TestFile.doc")

        extract_content_between_paragraphs(doc)
        extract_content_between_block_level_nodes(doc)
        extract_content_between_paragraph_styles(doc)
        extract_content_between_runs(doc)
        extract_content_using_field(doc)
        extract_content_between_bookmark(doc)
        extract_content_between_comment_range(doc)
    end

    def extract_content_between_paragraphs(doc)
        # Gather the nodes. The GetChild method uses 0-based index
        node_type = Rjb::import("com.aspose.words.NodeType")
        start_para = doc.getFirstSection().getChild(node_type.PARAGRAPH, 6, true)
        end_para =   doc.getFirstSection().getChild(node_type.PARAGRAPH, 10, true)
        
        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = extract_contents(start_para, end_para, true)
        
        # Insert the content into a new separate document and save it to disk.
        dst_doc = generate_document(doc, extracted_nodes)
        dst_doc.save(@data_dir + "TestFile.Paragraphs Out.doc")
    end    
    
    def extract_content_between_block_level_nodes(doc)
        # Gather the nodes. The GetChild method uses 0-based index
        node_type = Rjb::import("com.aspose.words.NodeType")
        start_para = doc.getLastSection().getChild(node_type.PARAGRAPH, 2, true)
        end_table = doc.getLastSection().getChild(node_type.TABLE, 0, true)
        
        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = extract_contents(start_para, end_table, true)

        # Lets reverse the array to make inserting the content back into the document easier.
        collections = Rjb::import("java.util.Collections")
        collections.reverse(extracted_nodes)
        
        while extracted_nodes.size() > 0 do
            # Insert the last node from the reversed list
            end_table.getParentNode().insertAfter(extracted_nodes.get(0), end_table)
            # Remove this node from the list after insertion.
            extracted_nodes.remove(0)
        end
        
        # Save the generated document to disk.
        doc.save(@data_dir + "TestFile.DuplicatedContent Out.doc")
    end    
    
    def extract_content_between_paragraph_styles(doc)
        # Gather a list of the paragraphs using the respective heading styles.
        paras_style_heading1 = paragraphs_by_style_name(doc, "Heading 1")
        paras_style_heading3 = paragraphs_by_style_name(doc, "Heading 3")
        
        # Use the first instance of the paragraphs with those styles.
        start_para1 = paras_style_heading1.get(0)
        end_para1 = paras_style_heading3.get(0)
        
        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extracted_nodes = extract_contents(start_para1, end_para1, false)

        # Insert the content into a new separate document and save it to disk.
        dst_doc = generate_document(doc, extracted_nodes)
        dst_doc.save(@data_dir + "TestFile.Styles Out.doc")
    end
    
    def extract_content_between_runs(doc)
        # Retrieve a paragraph from the first section.
        node_type = Rjb::import("com.aspose.words.NodeType")
        para = doc.getChild(node_type.PARAGRAPH, 7, true)
        
        # Use some runs for extraction.
        start_run = para.getRuns().get(1)
        end_run = para.getRuns().get(4)
        
        # Extract the content between these nodes in the document. Include these markers in the extraction.
        extracted_nodes = extract_contents(start_run, end_run, true)
        
        # Get the node from the list. There should only be one paragraph returned in the list.
        node = extracted_nodes.get(0)
        
        # Print the text of this node to the console.
        save_format = Rjb::import("com.aspose.words.SaveFormat")
        puts node.toString(save_format.TEXT)
    end
    
    def extract_content_using_field(doc)
        # Use a document builder to retrieve the field start of a merge field.
        builder = Rjb::import("com.aspose.words.DocumentBuilder").new(doc)
        
        # Pass the first boolean parameter to get the DocumentBuilder to move to the FieldStart of the field.
        # We could also get FieldStarts of a field using GetChildNode method as in the other examples.
        builder.moveToMergeField("Fullname", false, false)
        
        #/ The builder cursor should be positioned at the start of the field.
        node_type = Rjb::import("com.aspose.words.NodeType")
        start_field = builder.getCurrentNode()
        end_para = doc.getFirstSection().getChild(node_type.PARAGRAPH, 5, true)

        # Extract the content between these nodes in the document. Don't include these markers in the extraction.
        extracted_nodes = extract_contents(start_field, end_para, false)

        # Insert the content into a new separate document and save it to disk.
        dst_doc = generate_document(doc, extracted_nodes)
        dst_doc.save(@data_dir + "TestFile.Fields Out.doc")
    end
    
    def extract_content_between_bookmark(doc)
        # Retrieve the bookmark from the document.
        bookmark = doc.getRange().getBookmarks().get("Bookmark1")
        
        # We use the BookmarkStart and BookmarkEnd nodes as markers.
        bookmark_start = bookmark.getBookmarkStart()
        bookmark_end = bookmark.getBookmarkEnd()
        
        # Firstly extract the content between these nodes including the bookmark.
        extracted_nodes_inclusive = extract_contents(bookmark_start, bookmark_end, true)
        dst_doc = generate_document(doc, extracted_nodes_inclusive)
        dst_doc.save(@data_dir + "TestFile.BookmarkInclusive Out.doc")
        
        # Secondly extract the content between these nodes this time without including the bookmark.
        extracted_nodes_exclusive = extract_contents(bookmark_start, bookmark_end, false)
        dst_doc = generate_document(doc, extracted_nodes_exclusive)
        dst_doc.save(@data_dir + "TestFile.BookmarkExclusive Out.doc")
    end
    
    def extract_content_between_comment_range(doc)
        # This is a quick way of getting both comment nodes.
        # Your code should have a proper method of retrieving each corresponding start and end node.
        node_type = Rjb::import("com.aspose.words.NodeType")
        comment_start = doc.getChild(node_type.COMMENT_RANGE_START, 0, true)
        comment_end = doc.getChild(node_type.COMMENT_RANGE_END, 0, true)

        # Firstly extract the content between these nodes including the bookmark.
        extracted_nodes_inclusive = extract_contents(comment_start, comment_end, true)
        dst_doc = generate_document(doc, extracted_nodes_inclusive)
        dst_doc.save(@data_dir + "TestFile.CommentInclusive Out.doc")
        
        # Secondly extract the content between these nodes this time without including the bookmark.
        extracted_nodes_exclusive = extract_contents(comment_start, comment_end, false)
        dst_doc = generate_document(doc, extracted_nodes_exclusive)
        dst_doc.save(@data_dir + "TestFile.CommentExclusive Out.doc")
    end

=begin
    This is a method which extracts blocks of content from a document between specified nodes.
    
    Extracts a range of nodes from a document found between specified markers and returns a copy of those nodes. Content can be extracted
    between inline nodes, block level nodes, and also special nodes such as Comment or Boomarks. Any combination of different marker types can used.
    
    @param string startNode The node which defines where to start the extraction from the document. This node can be block or inline level of a body.
    @param string endNode The node which defines where to stop the extraction from the document. This node can be block or inline level of body.
    @param boolean isInclusive Should the marker nodes be included.
=end     
    def extract_contents(startNode, endNode, isInclusive)
        # First check that the nodes passed to this method are valid for use.
        verify_parameter_nodes(startNode, endNode)
        
        # Create a list to store the extracted nodes.
        nodes = Rjb::import("java.util.ArrayList").new
        
        # Keep a record of the original nodes passed to this method so we can split marker nodes if needed.
        originalStartNode = startNode
        originalEndNode = endNode
        
        # Extract content based on block level nodes (paragraphs and tables). Traverse through parent nodes to find them.
        # We will split the content of first and last nodes depending if the marker nodes are inline
        node_type = Rjb::import("com.aspose.words.NodeType")
        
        while (startNode.getParentNode().getNodeType() != node_type.BODY) do
            startNode = startNode.getParentNode()
        end

        while (endNode.getParentNode().getNodeType() != node_type.BODY) do
            endNode = endNode.getParentNode()
        end

        isExtracting = true
        isStartingNode = true
        isEndingNode = ''
        #The current node we are extracting from the document.
        currNode = startNode
        
        #Begin extracting content. Process all block level nodes and specifically split the first and last nodes when needed so paragraph formatting is retained.
        # Method is little more complex than a regular extractor as we need to factor in extracting using inline nodes, fields, bookmarks etc as to make it really useful.
        while (isExtracting) do
            # Clone the current node and its children to obtain a copy.
            cloneNode = currNode.deepClone(true)
            isEndingNode = currNode.equals(endNode)

            if (isStartingNode || isEndingNode) then
                # We need to process each marker separately so pass it off to a separate method instead.
                if (isStartingNode) then
                    process_marker(cloneNode, nodes, originalStartNode, isInclusive, isStartingNode, isEndingNode)
                    isStartingNode = false
                end
                # Conditional needs to be separate as the block level start and end markers maybe the same node.
                if (isEndingNode) then
                    process_marker(cloneNode, nodes, originalEndNode, isInclusive, isStartingNode, isEndingNode)
                    isExtracting = false
                end
            else
                # Node is not a start or end marker, simply add the copy to the list.
                nodes.add(cloneNode)
            end

            # Move to the next node and extract it. If next node is null that means the rest of the content is found in a different section.
            #if (currNode.getNextSibling() == null && isExtracting) then
            if ((currNode.getNextSibling()).nil? && isExtracting) then    
                # Move to the next section.
                nodeType = Rjb::import("com.aspose.words.NodeType")
                nextSection = currNode.getAncestor(nodeType.SECTION).getNextSibling()
                currNode = nextSection.getBody().getFirstChild()
            else
                # Move to the next node in the body.
                currNode = currNode.getNextSibling()
            end
        end
        # Return the nodes between the node markers.
        nodes
    end

=begin
    Checks the input parameters are correct and can be used. Throws an exception if there is any problem.
=end     
    def verify_parameter_nodes(startNode, endNode)
        # The order in which these checks are done is important.
        raise 'Start node cannot be null' if startNode.nil?
        raise 'End node cannot be null' if endNode.nil?
        raise "Start node and end node must belong to the same document" if (startNode.getDocument() == endNode.getDocument())
            
        nodeType = Rjb::import("com.aspose.words.NodeType")
        #raise "Start node and end node must be a child or descendant of a body" if (startNode.getAncestor(nodeType.BODY) == '' || endNode.getAncestor(nodeType.BODY) == '')
        raise "Start node and end node must be a child or descendant of a body" if (startNode.getAncestor(nodeType.BODY).nil? || endNode.getAncestor(nodeType.BODY).nil?)
            
        # Check the end node is after the start node in the DOM tree
        # First check if they are in different sections, then if they're not check their position in the body of the same section they are in.
        startSection = startNode.getAncestor(nodeType.SECTION)
        endSection = endNode.getAncestor(nodeType.SECTION)
        startIndex = startSection.getParentNode().indexOf(startSection)
        endIndex = endSection.getParentNode().indexOf(endSection)
        
        if (startIndex == endIndex) then
            raise "The end node must be after the start node in the body" if (startSection.getBody().indexOf(startNode) > endSection.getBody().indexOf(endNode))
        elsif (startIndex > endIndex) then
            raise "The section of end node must be after the section start node"
        end    
    end

    def generate_document(src_doc, nodes)
        # Create a blank document.
        dst_doc = Rjb::import("com.aspose.words.Document").new
        
        # Remove the first paragraph from the empty document.
        dst_doc.getFirstSection().getBody().removeAllChildren()
        
        # Import each node from the list into the new document. Keep the original formatting of the node.
        import_format_mode = Rjb::import("com.aspose.words.ImportFormatMode")
        importer = Rjb::import("com.aspose.words.NodeImporter").new(src_doc, dst_doc, import_format_mode.KEEP_SOURCE_FORMATTING)
        
        i = 0
        while i < nodes.size
            node = nodes.get(i)
            import_node = importer.importNode(node, true)
            dst_doc.getFirstSection().getBody().appendChild(import_node)
            i +=1
        end
        
        # Return the generated document.
        dst_doc
    end

    def process_marker(cloneNode, nodes, node, isInclusive, isStartMarker, isEndMarker)
        # If we are dealing with a block level node just see if it should be included and add it to the list.
        if (!is_inline(node)) then
            # Don't add the node twice if the markers are the same node
            if(!(isStartMarker && isEndMarker)) then
                if (isInclusive) then
                    nodes.add(cloneNode)
                end
            end
            return
        end

        # If a marker is a FieldStart node check if it's to be included or not.
        # We assume for simplicity that the FieldStart and FieldEnd appear in the same paragraph.
        nodeType = Rjb::import("com.aspose.words.NodeType")
        if (node.getNodeType() == nodeType.FIELD_START) then
            # If the marker is a start node and is not be included then skip to the end of the field.
            # If the marker is an end node and it is to be included then move to the end field so the field will not be removed.
            #if ((isStartMarker && !isInclusive) || (!isStartMarker && isInclusive)) then
            if ((isStartMarker && isInclusive.nil?) || (!isStartMarker && isInclusive)) then    
                #while (node.getNextSibling() != null && node.getNodeType() != nodeType.FIELD_END) do
                while (node.getNextSibling().nil? && (node.getNodeType() != nodeType.FIELD_END)) do    
                    node = node.getNextSibling()
                end    
            end
        end

        # If either marker is part of a comment then to include the comment itself we need to move the pointer forward to the Comment
        # node found after the CommentRangeEnd node.
        if (node.getNodeType() == nodeType.COMMENT_RANGE_END) then
            while (node.getNextSibling().nil? && (node.getNodeType() != nodeType.COMMENT)) do    
                node = node.getNextSibling()
            end    
        end

        # Find the corresponding node in our cloned node by index and return it.
        # If the start and end node are the same some child nodes might already have been removed. Subtract the
        # difference to get the right index.
        indexDiff = (node.getParentNode().getChildNodes().getCount() - cloneNode.getChildNodes().getCount())
        
        # Child node count identical.
        if (indexDiff == 0) then
            node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node))
        else
            node = cloneNode.getChildNodes().get(node.getParentNode().indexOf(node) - indexDiff)
        end
            
        # Remove the nodes up to/from the marker.
        isSkip = ''
        isProcessing = true
        isRemoving = isStartMarker
        nextNode = cloneNode.getFirstChild()
        #while (isProcessing && nextNode != null) do
        unless (isProcessing && nextNode.nil?)
            currentNode = nextNode
            isSkip = false
            if (currentNode == node) then
                if (isStartMarker) then
                    isProcessing = false
                    if isInclusive then
                        isRemoving = false
                    end        
                else
                    isRemoving = true
                    if isInclusive then
                        isSkip = true
                    end    
                end
            end
            nextNode = nextNode.getNextSibling()
            #if (isRemoving && !isSkip) then
            if (isRemoving && isSkip==false) then    
                currentNode.remove()
            end    
        end

        # After processing the composite node may become empty. If it has don't include it.
        if (!(isStartMarker && isEndMarker)) then
            if cloneNode.hasChildNodes() then
                nodes.add(cloneNode)
            end
        end
    end

    def is_inline(node)
        # Test if the node is desendant of a Paragraph or Table node and also is not a paragraph or a table a paragraph inside a comment class which is decesant of a pararaph is possible.
        node_type = Rjb::import("com.aspose.words.NodeType")
        #return ((node.getAncestor(node_type.PARAGRAPH) != null) || (node.getAncestor(node_type.TABLE) != null) && !(node.getNodeType() == nodeType.PARAGRAPH) || (node.getNodeType() == nodeType.TABLE))
        return ((node.getAncestor(node_type.PARAGRAPH).nil?) || (node.getAncestor(node_type.TABLE).nil?) && !(node.getNodeType() == node_type.PARAGRAPH) || (node.getNodeType() == node_type.TABLE))
    end

    def paragraphs_by_style_name(doc, style_name)
        # Create an array to collect paragraphs of the specified style.
        paragraphsWithStyle = Rjb::import("java.util.ArrayList").new
        
        # Get all paragraphs from the document.
        node_type = Rjb::import("com.aspose.words.NodeType")
        paragraphs = doc.getChildNodes(node_type.PARAGRAPH, true)
        paragraphs_count = paragraphs.getCount()
        #paragraphs_count = java_values($paragraphs_count)
        
        # Look through all paragraphs to find those with the specified style.
        i = 0
        while (i < paragraphs_count) do
            paragraphs = doc.getChildNodes(node_type.PARAGRAPH, true)
            paragraph = paragraphs.get(i)
            if (paragraph.getParagraphFormat().getStyle().getName() == style_name) then
                paragraphsWithStyle.add(paragraph)
            end
            i = i + 1
        end
        paragraphsWithStyle
    end

  end
end
