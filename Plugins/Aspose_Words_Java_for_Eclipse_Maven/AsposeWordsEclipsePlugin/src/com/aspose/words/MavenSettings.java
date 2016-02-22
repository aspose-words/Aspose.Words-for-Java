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
package com.aspose.words;

import org.eclipse.core.runtime.preferences.IEclipsePreferences;
import org.eclipse.core.runtime.preferences.InstanceScope;

/**
 *
 * @author Adeel
 */
public final class MavenSettings {

    private static final String PROP_LAST_ARCHETYPE_GROUPID = "lastArchetypeGroupId";
    private static final String PROP_LAST_ARCHETYPE_VERSION = "lastArchetypeVersion";

    private IEclipsePreferences getPreferences() {  	    	
    	return InstanceScope.INSTANCE.getNode(Activator.PLUGIN_ID);    	    	
    }

    /**
     *
     * @return
     */
    public String getLastArchetypeGroupId() {
        return getPreferences().get(PROP_LAST_ARCHETYPE_GROUPID, "com.mycompany");
    }

    /**
     *
     * @return
     */
    public String getLastArchetypeVersion() {
        return getPreferences().get(PROP_LAST_ARCHETYPE_VERSION, "1.0-SNAPSHOT");
    }

    /**
     *
     * @param version
     */
    public void setLastArchetypeVersion(String version) {
        putProperty(PROP_LAST_ARCHETYPE_VERSION, version);
    }

    /**
     *
     * @param groupId
     */
    public void setLastArchetypeGroupId(String groupId) {
        putProperty(PROP_LAST_ARCHETYPE_GROUPID, groupId);
    }

    private String putProperty(String key, String value) {
        String retval = getProperty(key);
        if (value != null) {
            getPreferences().put(key, value);
        } else {
            getPreferences().remove(key);
        }
        return retval;
    }
    private static final MavenSettings INSTANCE = new MavenSettings();

    /**
     *
     * @return
     */
    public static MavenSettings getDefault() {
        return INSTANCE;
    }

    private String getProperty(String key) {
        return getPreferences().get(key, null);
    }

}
