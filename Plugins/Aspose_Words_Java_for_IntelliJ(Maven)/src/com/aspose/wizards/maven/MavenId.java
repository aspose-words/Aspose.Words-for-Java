
package com.aspose.wizards.maven;

import org.jetbrains.annotations.Nullable;

public class MavenId {

    @Nullable
    private final String myGroupId;
    @Nullable
    private final String myArtifactId;
    @Nullable
    private final String myVersion;

    public MavenId(@Nullable String groupId, @Nullable String artifactId, @Nullable String version) {
        myGroupId = groupId;
        myArtifactId = artifactId;
        myVersion = version;
    }

    @Nullable
    public String getGroupId() {
        return myGroupId;
    }

    @Nullable
    public String getArtifactId() {
        return myArtifactId;
    }

    @Nullable
    public String getVersion() {
        return myVersion;
    }


}
