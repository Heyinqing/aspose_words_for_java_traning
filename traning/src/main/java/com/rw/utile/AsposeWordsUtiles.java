package com.rw.utile;

import com.aspose.words.License;

import java.io.InputStream;

public class AsposeWordsUtiles {
    public boolean GetLicense() {
        boolean result = false;
        try {
            InputStream license = AsposeWordsUtiles.class.getClassLoader().getResourceAsStream(
                    "\\license.xml"); // license路径
            License aposeLic = new License();
            aposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}
