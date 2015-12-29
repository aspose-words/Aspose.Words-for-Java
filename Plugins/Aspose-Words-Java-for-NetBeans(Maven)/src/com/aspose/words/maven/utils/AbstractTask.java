package com.aspose.words.maven.utils;

import org.netbeans.api.progress.aggregate.AggregateProgressFactory;
import org.netbeans.api.progress.aggregate.ProgressContributor;

/**
 * @author Adeel Ilyas <adeel.ilyas@aspose.com>
 */
public abstract class AbstractTask extends Thread {

    /**
     *
     */
    protected ProgressContributor p = null;

    /**
     *
     * @param id
     */
    public AbstractTask(String id) {
        p = AggregateProgressFactory.createProgressContributor(id);
    }

    /**
     *
     * @return
     */
    public ProgressContributor getProgressContributor() {
        return p;
    }
}
