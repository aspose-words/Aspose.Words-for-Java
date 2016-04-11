/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2016 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */

package com.aspose.dotcms.words;

import javax.servlet.ServletException;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import java.io.*;
import java.net.URL;

public class ExportToWordServlet extends HttpServlet {

    private static final long serialVersionUID = 42L;

    public ExportToWordServlet ( ) {
    }
    
    protected void doGet ( HttpServletRequest httpServletRequest, HttpServletResponse httpServletResponse) throws ServletException, IOException {
        //Get Web page URL
        String pageURL = getPageURL(httpServletRequest);
        
		try {
			//Save Web page content in Word Processing document
			String fileName = savePageContentInWordProcessingDocument(pageURL);
			//Send Document to Client
	        sendDocumentToClient(fileName, httpServletResponse);
		} catch (Exception e) {
			OutputStream os= httpServletResponse.getOutputStream();
	        os.write(Constants.HTML_TO_WORD_CONVERSION_ERROR_MESSAGE.getBytes());
			os.flush();
		}
    }
    
    protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
        doGet(request, response);
    }
    
    public String getPageURL(HttpServletRequest httpServletRequest) {
    	String pageURL = null;
    	
    	String scheme = httpServletRequest.getScheme();
    	String host = httpServletRequest.getHeader("Host");
        String path = httpServletRequest.getParameter("page_url");
        
        pageURL = scheme + "://" + host + path;
        
        return pageURL;
    }
    
    public String savePageContentInWordProcessingDocument(String pageURL) throws Exception {
    	String fileName = null;
    	
        URL url = new URL(pageURL);
        InputStream stream = url.openStream();
        // Load the entire document into memory
        Document doc = new Document(stream);
        // Save the document DOCX file format
        fileName = Constants.WORD_FILE_NAME;
        File file = new File(fileName);
        if(file.exists()) {
        	file.delete();
        }
        doc.save(fileName, SaveFormat.DOCX);
        stream.close();       
            
        return fileName;
    }
    
    public void sendDocumentToClient(String fileName, HttpServletResponse response) throws ServletException, IOException {
    	ServletOutputStream stream = null;
        BufferedInputStream buf = null;
        try {
        	
            stream = response.getOutputStream();
            File file = new File(fileName);
            
            response.setContentType("application/msword");
            response.addHeader("Content-Disposition", "attachment; filename="+ fileName);
            response.setContentLength((int) file.length());
	        
            FileInputStream input = new FileInputStream(file);
	        buf = new BufferedInputStream(input);
	        int readBytes = 0;
	        while ((readBytes = buf.read()) != -1) {
	        	stream.write(readBytes);
	        }
        } catch (IOException ioe) {
        	throw new ServletException(ioe.getMessage());
        } finally {
          if (stream != null)
        	  stream.close();
          if (buf != null)
            buf.close();
        }
    }
}