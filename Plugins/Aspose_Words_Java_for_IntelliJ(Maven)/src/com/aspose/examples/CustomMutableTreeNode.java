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

package com.aspose.examples;

import javax.swing.tree.DefaultMutableTreeNode;

/**
 *
 * @author Adeel Ilyas
 */
public final class CustomMutableTreeNode extends DefaultMutableTreeNode{

    private boolean folder;
    private String _ExPath;

    public void setTopTreeNodeText(String data) {
        this.userObject = data;
    }
    public CustomMutableTreeNode(String data)
    {
      this.userObject = data;  
    }

    /**
     * @return the _ExPath
     */
    public String getExPath() {
        return _ExPath;
    }

    /**
     * @param ExPath the _ExPath to set
     */
    public void setExPath(String ExPath) {
        this._ExPath = ExPath;
    }

    public boolean isFolder() {
        return folder;
    }

    public void setFolder(boolean folder) {
        this.folder = folder;
    }
}
