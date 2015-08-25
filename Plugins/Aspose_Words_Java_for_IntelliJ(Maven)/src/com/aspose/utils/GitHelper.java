/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package com.aspose.utils;

//import com.aspose.examples.otherexamples.OtherExamplesManager;

import org.eclipse.jgit.api.Git;
import org.eclipse.jgit.internal.storage.file.FileRepository;
import org.eclipse.jgit.lib.Repository;

import java.io.File;

/**
 * @author Adeel Ilyas
 * 
 */
public class GitHelper {
    public static void updateRepository(String localPath, String remotePath) throws Exception {
        Repository localRepo;
        try {
            localRepo = new FileRepository(localPath + "/.git");

            Git git = new Git(localRepo);
            {
                AsposeConstants.println("Cloning Repository [" + remotePath + "]....");
            }

            // First try to clone the repository
            try {
                Git.cloneRepository().setURI(remotePath).setDirectory(new File(localPath)).call();
            } catch (Exception ex) {
                // If clone fails, try to pull the changes
                try {
                    git.pull().call();
                } catch (Exception exPull) {
                    // Pull also failed. Throw this exception to caller
                    {
                        AsposeConstants.println("Pull also failed.");
                    }
                    throw exPull; // throw it
                }
            }
        } catch (Exception ex) {
            throw new Exception("Could not download Repository from Github. Error: " + ex.getMessage());
        }
    }

    public static void syncRepository(String localPath, String remotePath) throws Exception {
        Repository localRepo;
        try {
            localRepo = new FileRepository(localPath + "/.git");

            Git git = new Git(localRepo);

                AsposeConstants.println("Syncronizing Repository [" + remotePath + "]....");
               // Pull the changes
            try {
                git.pull().call();
            } catch (Exception exPull) {
                // If pull failed. Throw this exception to caller
                {
                    AsposeConstants.println("Pull failed.");
                }
                throw exPull; // throw it
            }

        } catch (Exception ex) {
            throw new Exception("Could not update Repository from Github. Error: " + ex.getMessage());
        }
    }


}
