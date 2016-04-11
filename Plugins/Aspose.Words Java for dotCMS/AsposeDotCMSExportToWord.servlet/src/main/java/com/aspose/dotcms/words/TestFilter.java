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

import javax.servlet.*;
import javax.servlet.http.HttpServletRequest;
import java.io.IOException;

public class TestFilter implements Filter {

    private final String name;

    public TestFilter ( String name ) {
        this.name = name;
    }

    public void init ( FilterConfig config ) throws ServletException {

        doLog( "Init with config [" + config + "]" );
    }

    public void doFilter ( ServletRequest req, ServletResponse res, FilterChain chain ) throws IOException, ServletException {

        if ( req instanceof HttpServletRequest ) {
            doLog( "Filter request [" + ((HttpServletRequest) req).getRequestURI() + "]" );
        } else {
            doLog( "Filter request [" + req + "]" );
        }

        chain.doFilter( req, res );
    }

    public void destroy () {
        doLog( "Destroyed filter" );
    }

    private void doLog ( String message ) {
        System.out.println( "## [" + this.name + "] " + message );
    }

}