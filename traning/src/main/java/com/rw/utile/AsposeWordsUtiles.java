package com.rw.utile;

import com.aspose.words.License;

import java.io.InputStream;

public class AsposeWordsUtiles {
    public static boolean GetLicense() {
        boolean result = false;
        try {
            InputStream license = AsposeWordsUtiles.class.getClassLoader().getResourceAsStream(
                    "E:\\project\\aspose_words_for_java_traning\\traning\\src\\main\\resources\\template\\license.xml"); // license路径
            License aposeLic = new License();
            aposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }
}
