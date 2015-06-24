module Asposewordsjavaforruby
  module ProcessComments
    def initialize()
        # The path to the documents directory.
        data_dir = File.dirname(File.dirname(File.dirname(__FILE__))) + '/data/'
        
        # Open the document.
        doc = Rjb::import('com.aspose.words.Document').new(data_dir + 'TestComments.doc')

        # Get all comments from document
        extract_comments(doc)
              
        # Remove comments by the "pm" author.
        remove_comment($doc, "pm");
        puts "Comments from 'pm' are removed!"
        
        # Remove all comments.
        remove_comments(doc)
        puts "All comments are removed!"

        # Save the document.
        doc.save(data_dir + "TestComments Out.doc")    
    end    

    def extract_comments(doc)
        # Call method
        collected_comments  = Rjb::import('java.util.ArrayList').new

        # Collect all comments in the document
        node_type = Rjb::import('com.aspose.words.NodeType')
        comments = doc.getChildNodes(node_type.COMMENT, true).toArray()
        
        save_format = Rjb::import('com.aspose.words.SaveFormat')
        
        comments.each do |comment|
            author = comment.getAuthor()
            date_time = comment.getDateTime()
            format = comment.toString(save_format.TEXT)
            puts "Author:" + author.to_s + " DateTime:" + date_time.to_string + " Comment:" + format.to_s
        end 
    end

    def remove_comment(doc, author_name)
        raise 'author_name not specified.' if author_name.empty?

        # Collect all comments in the document
        node_type = Rjb::import('com.aspose.words.NodeType')
        comments = doc.getChildNodes(node_type.COMMENT, true)
        comments_count = comments.getCount()
        
        # Look through all comments and remove those written by the authorName author.
        i = comments_count
        i = i - 1
        while (i >= 0) do
            comment = comments.get(i)
            author = comment.getAuthor().chomp('"').reverse.chomp('"').reverse
            if (author == author_name) then
                comment.remove()
            end
            i = i - 1
        end
    end

    def remove_comments(doc)
        # Collect all comments in the document
        node_type = Rjb::import('com.aspose.words.NodeType')
        comments = doc.getChildNodes(node_type.COMMENT, true)
        comments.clear()
    end

  end
end
