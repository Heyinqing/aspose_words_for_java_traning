package com.rw.utile;

import com.aspose.words.License;

import java.io.FileInputStream;
import java.io.InputStream;

public class AsposeWordsUtiles {
    public static boolean GetLicense() {
        boolean result = false;
        try {
//            FileInputStream fileInputStream = new FileInputStream("E:\\Desktop\\license.xml");

            InputStream license = AsposeWordsUtiles.class.getResourceAsStream(
                    "license.xml"); // license路径
            License aposeLic = new License();
            aposeLic.setLicense(license);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            return result;
        }
    }
}
