
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


package com.aspose.utils.execution;

import com.intellij.openapi.application.ApplicationManager;
import com.intellij.openapi.command.CommandProcessor;
import com.intellij.openapi.project.Project;
import com.intellij.openapi.startup.StartupManager;

/**
 * @author Adeel Ilyas <adeel.ilyas@aspose.com>
 */

public class RunnableHelper

{
    public static void runWhenInitialized(final Project project, final Runnable r) {
        if (project.isDisposed()) return;

        if (!project.isInitialized()) {
            StartupManager.getInstance(project).registerStartupActivity(new RunnableHelper.WriteAction(r));
            return;
        }

        RunnableHelper.runWriteCommand(project, r);

    }

    public static void runReadCommand(Project project, Runnable cmd)

    {

        CommandProcessor.getInstance().executeCommand(project, new ReadAction(cmd), "Aspose", "Components");

    }


    public static void runWriteCommand(Project project, Runnable cmd)

    {

        CommandProcessor.getInstance().executeCommand(project, new WriteAction(cmd), "Aspose", "Components");

    }


    public static class ReadAction implements Runnable

    {

        public ReadAction(Runnable cmd)

        {

            this.cmd = cmd;

        }


        public void run()

        {

            ApplicationManager.getApplication().runReadAction(cmd);

        }


        Runnable cmd;

    }


    public static class WriteAction implements Runnable

    {

        public WriteAction(Runnable cmd)

        {

            this.cmd = cmd;

        }


        public void run()

        {
            ApplicationManager.getApplication().invokeLater(new Runnable() {
                public void run() {
                    ApplicationManager.getApplication().runWriteAction(cmd);
                }
            });
        }


        Runnable cmd;

    }


    private RunnableHelper() {
    }

}