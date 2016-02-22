/*
 * The MIT License (MIT)
 *
 * Copyright (c) 1998-2015 Aspose Pty Ltd.
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package com.aspose.words.maven.utils;

/*
 * @author Adeel Ilyas <adeel.ilyas@aspose.com>
 *
 */
// Singleton Class

/**
 *
 * @author Adeel
 */
public class AsposeWordsJavaAPI extends AsposeJavaAPI {

    private final String _name = AsposeConstants.API_NAME;
    private final String _mavenRepositoryURL = "http://maven.aspose.com/repository/ext-release-local/com/aspose/aspose-words/";
    private final String _remoteExamplesRepository = "https://github.com/asposewords/Aspose_Words_Java";

    /**
     * @return the _name
     */
    @Override
    public String get_name() {
        return _name;
    }

    /**
     * @return the _mavenRepositoryURL
     */
    @Override
    public String get_mavenRepositoryURL() {
        return _mavenRepositoryURL;
    }

    /**
     * @return the _remoteExamplesRepository
     */
    @Override
    public String get_remoteExamplesRepository() {
        return _remoteExamplesRepository;
    }

    // Singleton instance
    private static AsposeJavaAPI asposeWordsAPI;

    /**
     *
     * @return
     */
    public static AsposeJavaAPI getInstance() {
        return asposeWordsAPI;
    }

    /**
     *
     * @param asposeMavenProjectManager
     * @return
     */
    public static AsposeJavaAPI initialize(AsposeMavenProjectManager asposeMavenProjectManager) {
        asposeWordsAPI = new AsposeWordsJavaAPI();
        asposeWordsAPI.asposeMavenProjectManager = asposeMavenProjectManager;
        return asposeWordsAPI;
    }

    private AsposeWordsJavaAPI() {
    }
}
