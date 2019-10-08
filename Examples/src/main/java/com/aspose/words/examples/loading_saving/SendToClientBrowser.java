package com.aspose.words.examples.loading_saving;

public class SendToClientBrowser {
    public static void main(String[] args) throws Exception {
        /*
        //To send a document to Client browser, you need to implement HTTPServlet in your web project
        // Please check following code snippet, called from the web-app index page (because the POST method is chosen for the input form).
        //ExStart:SendToClientBrowser

        protected void doPost(HttpServletRequest request,HttpServletResponse response) throws ServletException, IOException {
            if (request.getParameter("button") != null) {

                // Get the output format selected by the user.
                String formatType = "PDF";
                Boolean openNewWindow = false;
                try {
                    com.aspose.words.Document doc = new com.aspose.words.Document(MyDir+"Test File.docx");

                    String fileName = "outDocument"+"."+ formatType;
                    response.setContentType("application/pdf");
                    // Add the Response header
                    if (openNewWindow)
                        response.setHeader("content-disposition","attachment; filename=" + fileName);
                    else
                        response.addHeader("content-disposition","inline; filename=" + fileName);

                    doc.save(response.getOutputStream(),com.aspose.words.SaveFormat.PDF);

                    response.flushBuffer();

                    System.out.println("Process Completed Successfully");
                } catch (Exception e) {
                    throw new RuntimeException("Process failed: " + e.getMessage());
                }
            }
        }
		//ExEnd:SendToClientBrowser
		*/

    }

}
