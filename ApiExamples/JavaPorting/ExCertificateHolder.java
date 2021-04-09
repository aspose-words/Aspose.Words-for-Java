package ApiExamples;

// ********* THIS FILE IS AUTO PORTED *********

import org.testng.annotations.Test;
import com.aspose.ms.System.IO.File;
import com.aspose.words.CertificateHolder;
import com.aspose.ms.System.IO.FileStream;
import com.aspose.ms.System.IO.FileMode;
import org.bouncycastle.jcajce.provider.keystore.pkcs12.PKCS12KeyStoreSpi;
import java.util.Iterator;
import com.aspose.ms.System.msConsole;

@Test
public class ExCertificateHolder extends ApiExampleBase
{
    @Test
    public void create() throws Exception
    {
        //ExStart
        //ExFor:CertificateHolder.Create(Byte[], SecureString)
        //ExFor:CertificateHolder.Create(Byte[], String)
        //ExFor:CertificateHolder.Create(String, String, String)
        //ExSummary:Shows how to create CertificateHolder objects.
        // Below are four ways of creating CertificateHolder objects.
        // 1 -  Load a PKCS #12 file into a byte array and apply its password:
        byte[] certBytes = File.readAllBytes(getMyDir() + "morzal.pfx");
        CertificateHolder.create(certBytes, "aw");

        // 2 -  Load a PKCS #12 file into a byte array, and apply a secure password:
        SecureString password = new NetworkCredential("", "aw").SecurePassword;
        // JAVA-deleted Create(): Java hasn't SecureString analog: 1) it should be low-level-platform-dependent, but 2) can't be absolutely safe.

        // If the certificate has private keys corresponding to aliases,
        // we can use the aliases to fetch their respective keys. First, we will check for valid aliases.
        FileStream certStream = new FileStream(getMyDir() + "morzal.pfx", FileMode.OPEN);
        try /*JAVA: was using*/
        {
            PKCS12KeyStoreSpi.BCPKCS12KeyStore pkcs12Store = new PKCS12KeyStoreSpi.BCPKCS12KeyStore(certStream, "aw".toCharArray());
            Iterator enumerator = pkcs12Store.getAliases().iterator();

            while (enumerator.hasNext())
            {
                if (enumerator.next() != null)
                {
                    String currentAlias = enumerator.next().toString();
                    if (pkcs12Store.isKeyEntry(currentAlias) && pkcs12Store.getKey(currentAlias).Key.isPrivate())
                    {
                        System.out.println("Valid alias found: {enumerator.Current}");
                    }
                }
            }
        }
        finally { if (certStream != null) certStream.close(); }

        // 3 -  Use a valid alias:
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", "c20be521-11ea-4976-81ed-865fbbfc9f24");

        // 4 -  Pass "null" as the alias in order to use the first available alias that returns a private key:
        CertificateHolder.create(getMyDir() + "morzal.pfx", "aw", null);
        //ExEnd
    }
}
