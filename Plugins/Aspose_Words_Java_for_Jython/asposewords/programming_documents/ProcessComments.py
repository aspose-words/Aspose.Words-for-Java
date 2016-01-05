from asposewords import Settings
from com.aspose.words import Document
from com.aspose.words import NodeType
from com.aspose.words import SaveFormat

#from java.util import ArrayList

class ProcessComments:

    def __init__(self):
        dataDir = Settings.dataDir + 'programming_documents/'
        
        # Open the document.
        doc = Document(dataDir + "TestFile.doc")

        #ExStart
        #ExId:ProcessComments_Main
        #ExSummary: The demo-code that illustrates the methods for the comments extraction and removal.
        # Extract the information about the comments of all the authors.

        comments = self.extract_comments(doc)

        for comment in comments:
            print comment
            
        # Remove all comments.
        self.remove_comments(doc)
        print "All comments are removed!"

        # Save the document.
        doc.save(dataDir + "Comments.doc")
    
    def extract_comments(self, *args):
        doc = args[0]
        collectedComments = []
        # Collect all comments in the document
        comments = doc.getChildNodes(NodeType.COMMENT, True)
        
        # Look through all comments and gather information about them.
        for comment in comments :
            if 1 < len(args) and args[1] is not None :
                authorName = args[1]
                if str(comment.getAuthor()) == authorName:
                    collectedComments.append(str(comment.getAuthor()) + " " + str(comment.getDateTime()) + " " + comment.toString(SaveFormat.TEXT))
            else:
                collectedComments.append(str(comment.getAuthor()) + " " + str(comment.getDateTime()) + " " + comment.toString(SaveFormat.TEXT))

        return collectedComments
    
    def remove_comments(self,*args):

        doc = args[0]
        if 1 < len(args) and args[1] is not None :
                authorName = args[1]
        
        # Collect all comments in the document
        comments = doc.getChildNodes(NodeType.COMMENT, True)
        comments_count = comments.getCount()
        
        # Look through all comments and remove those written by the authorName author.
        i = comments_count
        i = i - 1
        while i >= 0 :
            comment = comments.get(i)
            if 1 < len(args) and args[1] is not None :
                authorName = args[1]
                if str(comment.getAuthor()) == authorName:
                    comment.remove()
            else:
                comment.remove()
            i = i - 1

if __name__ == '__main__':        
    ProcessComments()