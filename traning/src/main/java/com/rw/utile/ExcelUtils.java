package com.rw.utile;

import com.aspose.words.License;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * Aspose for Java 对Excel处理的工具类
 * <p>
 * Aspose的API文档请参考地址：https://apireference.aspose.com/java/cells/com.aspose.cells/Picture
 */
public class ExcelUtils {




    /**
     * license认证，保证license.xml放到resource根目录下
     *
     * @return
     */
    public boolean GetLicense() {
        boolean result = false;
        try {
            InputStream license = ExcelUtils.class.getClassLoader().getResourceAsStream(
                    "\\license.xml"); // license路径
            License aposeLic = new License();
            aposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

    /**
     * license认证，保证license.xml放到resource根目录下
     *
     * @return
     */
    public boolean GetWordLicense() {
        boolean result = false;
        try {
            InputStream license = ExcelUtils.class.getClassLoader().getResourceAsStream(
                    "license.xml"); // license路径
            com.aspose.words.License aposeLic = new com.aspose.words.License();
            aposeLic.setLicense(license);
            result = true;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return result;
    }

}
