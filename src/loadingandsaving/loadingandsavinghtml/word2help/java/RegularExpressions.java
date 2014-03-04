/* 
 * Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
 *
 * This file is part of Aspose.Words. The source code in this file
 * is only intended as a supplement to the documentation, and is provided
 * "as is", without warranty of any kind, either expressed or implied.
 */
 
package loadingandsaving.loadingandsavinghtml.word2help.java;

import java.util.regex.Pattern;


/**
 * Central storage for regular expressions used in the project.
 */
public class RegularExpressions
{
    // This class is static. No instance creation is allowed.
    private RegularExpressions() throws Exception {}

    /**
     * Regular expression specifying html title (framing tags excluded).
     */
    public static Pattern getHtmlTitle() throws Exception
    {
        if (gHtmlTitle == null)
        {
            gHtmlTitle = Pattern.compile(HTML_TITLE_PATTERN,
                Pattern.CASE_INSENSITIVE);
        }
        return gHtmlTitle;
    }

    /**
     * Regular expression specifying html head.
     */
    public static Pattern getHtmlHead() throws Exception
    {
        if (gHtmlHead == null)
        {
            gHtmlHead = Pattern.compile(HTML_HEAD_PATTERN,
                Pattern.CASE_INSENSITIVE);
        }
        return gHtmlHead;
    }

    /**
     * Regular expression specifying space right after div keyword in the first div declaration of html body.
     */
    public static Pattern getHtmlBodyDivStart() throws Exception
    {
        if (gHtmlBodyDivStart == null)
        {
            gHtmlBodyDivStart = Pattern.compile(HTML_BODY_DIV_START_PATTERN,
                Pattern.CASE_INSENSITIVE);
        }
        return gHtmlBodyDivStart;
    }

    private static final String HTML_TITLE_PATTERN = "(?<=\\<title\\>).*?(?=\\</title\\>)";
    private static Pattern gHtmlTitle;

    private static final String HTML_HEAD_PATTERN = "\\<head\\>.*?\\</head\\>";
    private static Pattern gHtmlHead;

    private static final String HTML_BODY_DIV_START_PATTERN = "(?<=\\<body\\>\\s{0,200}\\<div)\\s";
    private static Pattern gHtmlBodyDivStart;
}