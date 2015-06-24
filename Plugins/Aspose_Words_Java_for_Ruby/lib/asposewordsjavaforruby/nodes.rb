module Asposewordsjavaforruby
  module Nodes
    def initialize()
        # get nodes
        get_nodes()
    end
        
    def get_nodes()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/quickstart/'

        # Create a new document.
        doc = Rjb::import('com.aspose.words.Document').new()

        # Creates and adds a paragraph node to the document.
        para = Rjb::import("com.aspose.words.Paragraph").new(doc)

        # Typed access to the last section of the document.
        section = doc.getLastSection()
        section.getBody().appendChild(para)
        
        # Next print the node type of one of the nodes in the document.
        node_type = doc.getFirstSection().getBody().getNodeType()
        node = Rjb::import("com.aspose.words.Node")
        puts "NodeType: " + node.nodeTypeToString(node_type)
    end 

  end
end
