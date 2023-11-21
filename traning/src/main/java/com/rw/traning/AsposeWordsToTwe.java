package com.rw.traning;

import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;
import cn.hutool.json.JSONUtil;
import com.aspose.words.*;
import com.aspose.words.Font;
import com.rw.utile.ExcelUtils;

import java.awt.*;
import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.*;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * @author shangmi
 * @title AsposeWordsToTwe
 * @date 2023/8/30 10:28
 * @description TODO
 */
public class AsposeWordsToTwe {
    public static final Map<String, String> TITLE = new HashMap<String, String>() {{
        put("surveyTime", "调查时间：");
        put("feedbackTime", "反馈时间：");
        put("respondentType", "调查对象类型：");
        put("targetCompany", "调查对象公司：");
        put("baseName", "基地名称：");
        put("projectState", "项目状态：");
        put("projectName", "项目名称：");
    }};
    public static final List<String> MAPKEY = Arrays.asList("surveyTime", "feedbackTime", "respondentType", "targetCompany", "baseName", "projectState", "projectName");

    public static void main(String[] args) throws IOException {
        String data = data();

        JSONArray jsonArray = JSONUtil.parseArray(data);
        File file = new File("E:/testToOne.zip");
        // OutputStream out = Files.newOutputStream(Paths.get("E:/testToOne.zip"));
        ZipOutputStream zipOutputStream = new ZipOutputStream(new FileOutputStream(file));
        for (int i = 0; i < jsonArray.size(); i++) {

            ByteArrayOutputStream subject = subject(jsonArray.getStr(i));
            InputStream inputStream = new ByteArrayInputStream(subject.toByteArray());
            JSONObject jsonObject = jsonArray.getJSONObject(i);
            zipOutputStream.putNextEntry(
                    new ZipEntry(
                            jsonObject.getStr("projectName")
                                    + "-"
                                    + jsonObject.getStr("surveyTime").replace("-", "")
                                    + "-"
                                    + (i + 1)
                                    + ".docx")
            );
            // zipOutputStream.write(subject.toByteArray());
            int let = 0;
            byte[] bytes = new byte[1024];
            while ((let = inputStream.read(bytes))>0){
                zipOutputStream.write(bytes,0,let);
            }
            // inputStream.close();
        }
    }

    /**
     * 压缩多个文件到本地
     *
     * @param localFileName 本地文件名
     * @param files         File文件集合
     */
    public static void zipFiles(String localFileName, List<File> files) {
        ZipOutputStream zipOut = null;
        try {
            zipOut = new ZipOutputStream(new FileOutputStream(localFileName));


            for (File file : files) {
                FileInputStream fis = null;
                try {
                    fis = new FileInputStream(file);
                    //  单个文件名称
                    zipOut.putNextEntry(new ZipEntry(file.getName()));

                    //	输出文件
                    int len = 0;
                    byte[] buffer = new byte[1024];
                    while ((len = fis.read(buffer)) > 0) {
                        zipOut.write(buffer, 0, len);
                    }
                } catch (Exception e) {
                } finally {
                    if (fis != null) {
                        try {
                            fis.close();
                        } catch (IOException e) {
                        }
                    }
                }
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } finally {
            if (zipOut != null) {
                try {
                    zipOut.close();
                } catch (IOException e) {
                }
            }
        }
    }

    public static ByteArrayOutputStream subject(String data) {
        JSONObject jsonObject = JSONUtil.parseObj(data);
        AsposeWordsUtile asposeWordsUtile = new AsposeWordsUtile();
        try {
            Document nodes = new Document();
            DocumentBuilder builder = new DocumentBuilder(nodes);
            // 修改纸张样式
            PageSetup pageSetup = builder.getPageSetup();
            pageSetup.setPaperSize(PaperSize.A4);
            pageSetup.setOrientation(Orientation.PORTRAIT);
            pageSetup.setVerticalAlignment(PageVerticalAlignment.TOP);
            pageSetup.setLeftMargin(90);
            pageSetup.setTopMargin(72);
            pageSetup.setBottomMargin(72);
            pageSetup.setRightMargin(90);

            Date date = new Date();
            asposeWordsUtile.totalTitle("上海振华重工" + jsonObject.get("baseName") + "满意度调查问卷", builder);

            builder.writeln();
            builder.writeln();
            contentTable(jsonObject, builder);

            JSONArray jsonArray = jsonObject.getJSONArray("dataMap");
            jsonArray.forEach(info -> {
                builder.writeln();
                builder.writeln();
                JSONObject contents = JSONUtil.parseObj(info);
                subjectTitle(contents, builder);
                multipleChoiceOption(contents.getJSONArray("multipleChoice"), builder);
                builder.writeln();
                JSONArray checkBoxArray = contents.getJSONArray("checkBox");
                if (checkBoxArray.size() == 0) {
                    return;
                }
                asposeWordsUtile.threeLevelTitle("不满意原因：", builder);
                checkBoxOption(checkBoxArray, builder);
                builder.writeln();
                asposeWordsUtile.threeLevelTitle("其他原因：", builder);
                asposeWordsUtile.textPart(contents.getStr("rest"), builder);
            });
            ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
            nodes.save(byteArrayOutputStream, SaveFormat.DOCX);
            return byteArrayOutputStream;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }


    }

    /**
     * 表格内容生成
     *
     * @param contents
     * @param builder
     * @throws Exception
     */
    public static void contentTable(JSONObject contents, DocumentBuilder builder) throws Exception {
        // 设置样式
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //
        Font font = builder.getFont();

        paragraphFormat.clearFormatting();
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        font.clearFormatting();
        font.setSize(12);

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        for (int i = 0; i < MAPKEY.size(); i++) {
            if (i != 0 && i % 2 == 0) {
                builder.endRow();
            }
            String key = MAPKEY.get(i);

            builder.insertCell();
            font.setBold(true);
            builder.write(TITLE.get(key));
            builder.insertCell();
            font.setBold(false);
            builder.write(contents.get(key).toString());
            if ("projectName".equals(key)) {
                builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
                builder.insertCell();
                builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
                builder.insertCell();
                builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
            }
        }

        builder.endTable();
    }

    public static void subjectTitle(JSONObject contents, DocumentBuilder builder) {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        paragraphFormat.clearFormatting();

        paragraphFormat.setFirstLineIndent(32);

        // 对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        Font font = builder.getFont();
        font.setSize(16);
        font.setName("仿宋_GB2312");
        font.setColor(Color.BLACK);
        font.setBold(true);

        builder.writeln(contents.get("questionNumber") + "." + contents.get("problem"));
        font.clearFormatting();

    }

    public static void multipleChoiceOption(JSONArray contents, DocumentBuilder builder) {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        paragraphFormat.clearFormatting();

        paragraphFormat.setFirstLineIndent(32);

        // 对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        Font font = builder.getFont();
        font.setSize(16);
        font.setColor(Color.BLACK);
        try {
            for (int i = 0; i < contents.size(); i++) {
                JSONObject content = JSONUtil.parseObj(contents.get(i));
                builder.moveToMergeField("c1");
                font.setName("Wingdings 2");
                if (content.getBool("choose")) {
                    builder.write("\u0098");
                } else {
                    builder.write("\u0099");
                }
                font.setName("仿宋_GB2312");
                builder.writeln(content.getStr("option"));

            }
            font.clearFormatting();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    public static void checkBoxOption(JSONArray contents, DocumentBuilder builder) {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        paragraphFormat.clearFormatting();

        paragraphFormat.setFirstLineIndent(32);

        // 对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        Font font = builder.getFont();
        font.setSize(16);
        font.setColor(Color.BLACK);
        try {
            for (int i = 0; i < contents.size(); i++) {
                JSONObject content = JSONUtil.parseObj(contents.get(i));
                builder.moveToMergeField("c1");
                font.setName("Wingdings 2");
                if (content.getBool("choose")) {
                    builder.write("\u0052");
                } else {
                    builder.write("\u00A3");
                }
                font.setName("仿宋_GB2312");
                builder.writeln(content.getStr("option"));

            }

            font.clearFormatting();
        } catch (Exception e) {
            e.printStackTrace();
        }


    }


    public static String data() {
        return "[{\n" +
                "\t\"feedbackTime\": \"2023-08-30 16:36:09\",\n" +
                "\t\"surveyTime\": \"2023-08-30\",\n" +
                "\t\"targetCompany\": \"上海远东工程咨询监理\",\n" +
                "\t\"dataMap\": [{\n" +
                "\t\t\"rest\": \"you dou hao\",\n" +
                "\t\t\"problem\": \"问题1\",\n" +
                "\t\t\"checkBox\": [{\n" +
                "\t\t\t\"choose\": true,\n" +
                "\t\t\t\"option\": \"涉及船东义务的甲板准备未实施\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"现场主管/项目经理外语沟通不足\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"整改未结束进入下道工序，有闯关现象\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": true,\n" +
                "\t\t\t\"option\": \"未提供合适建议\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"开工准备不充分\"\n" +
                "\t\t}],\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"一般\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": true,\n" +
                "\t\t\t\"option\": \"不满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常不满意\"\n" +
                "\t\t}],\n" +
                "\t\t\"questionNumber\": \"1\"\n" +
                "\t}, {\n" +
                "\t\t\"rest\": \"汉字不满意\",\n" +
                "\t\t\"problem\": \"问题2\",\n" +
                "\t\t\"checkBox\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"设计图纸不符合规范或项目技术要求\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"设计图纸不完整、图纸质量较差，修改较多\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"提出的设计修改意见未能及时落实\"\n" +
                "\t\t}],\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"一般\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"不满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": true,\n" +
                "\t\t\t\"option\": \"非常不满意\"\n" +
                "\t\t}],\n" +
                "\t\t\"questionNumber\": \"2\"\n" +
                "\t}],\n" +
                "\t\"respondentType\": \"监理\",\n" +
                "\t\"projectName\": \"和黄墨西哥EIT码头岸桥\",\n" +
                "\t\"baseName\": \"长兴分公司\",\n" +
                "\t\"projectState\": \"经营\"\n" +
                "}, {\n" +
                "\t\"feedbackTime\": \"\",\n" +
                "\t\"surveyTime\": \"2023-08-30\",\n" +
                "\t\"targetCompany\": \"振华\",\n" +
                "\t\"dataMap\": [{\n" +
                "\t\t\"rest\": \"\",\n" +
                "\t\t\"problem\": \"问题1\",\n" +
                "\t\t\"checkBox\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"涉及船东义务的甲板准备未实施\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"现场主管/项目经理外语沟通不足\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"整改未结束进入下道工序，有闯关现象\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"未提供合适建议\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"开工准备不充分\"\n" +
                "\t\t}],\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"一般\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"不满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常不满意\"\n" +
                "\t\t}],\n" +
                "\t\t\"questionNumber\": \"1\"\n" +
                "\t}, {\n" +
                "\t\t\"rest\": \"\",\n" +
                "\t\t\"problem\": \"问题2\",\n" +
                "\t\t\"checkBox\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"设计图纸不符合规范或项目技术要求\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"设计图纸不完整、图纸质量较差，修改较多\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"提出的设计修改意见未能及时落实\"\n" +
                "\t\t}],\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"一般\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"不满意\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"choose\": false,\n" +
                "\t\t\t\"option\": \"非常不满意\"\n" +
                "\t\t}],\n" +
                "\t\t\"questionNumber\": \"2\"\n" +
                "\t}],\n" +
                "\t\"respondentType\": \"内部用户\",\n" +
                "\t\"projectName\": \"和黄墨西哥EIT码头岸桥\",\n" +
                "\t\"baseName\": \"长兴分公司\",\n" +
                "\t\"projectState\": \"经营\"\n" +
                "}]";
    }
}
