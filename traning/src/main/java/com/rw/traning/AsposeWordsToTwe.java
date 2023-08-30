package com.rw.traning;

import cn.hutool.json.JSONArray;
import cn.hutool.json.JSONObject;
import cn.hutool.json.JSONUtil;
import com.aspose.words.*;
import com.aspose.words.Font;
import com.rw.utile.ExcelUtils;

import java.awt.*;
import java.util.*;
import java.util.List;

/**
 * @author shangmi
 * @title AsposeWordsToTwe
 * @date 2023/8/30 10:28
 * @description TODO
 */
public class AsposeWordsToTwe {
    public static final Map<String,String> TITLE = new HashMap<String,String>(){{
       put("surveyTime","调查时间：");
       put("feedbackTime","反馈时间：");
       put("respondentType","调查对象类型：");
       put("targetCompany","调查对象公司：");
       put("baseName","基地名称：");
       put("projectState","项目状态：");
       put("projectName","项目名称：");
    }};
    public static final List<String> MAPKEY = Arrays.asList("surveyTime","feedbackTime","respondentType","targetCompany","baseName","projectState","projectName");

    public static void main(String[] args) throws Exception {
        String data = data();
        JSONObject jsonObject = JSONUtil.parseObj(data);
        AsposeWordsUtile asposeWordsUtile = new AsposeWordsUtile();
        Document nodes = new Document();
        DocumentBuilder builder = new DocumentBuilder(nodes);
        //修改纸张样式
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setPaperSize(PaperSize.A4);
        pageSetup.setOrientation(Orientation.PORTRAIT);
        pageSetup.setVerticalAlignment(PageVerticalAlignment.TOP);
        pageSetup.setLeftMargin(90);
        pageSetup.setTopMargin(72);
        pageSetup.setBottomMargin(72 );
        pageSetup.setRightMargin(90);

        Date date = new Date();
        asposeWordsUtile.totalTitle("上海振华重工"+jsonObject.get("baseName")+"满意度调查问卷",builder);
        contentTable(jsonObject,builder);

        JSONArray jsonArray = jsonObject.getJSONArray("subject");
        jsonArray.forEach(contexts ->{
            subjectTitle((JSONObject) contexts,builder);
        });

    }

    /**
     * 表格内容生成
     * @param contents
     * @param builder
     * @throws Exception
     */
    private static void contentTable(JSONObject contents, DocumentBuilder builder) throws Exception {
        //设置样式
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //
        Font font = builder.getFont();

        paragraphFormat.clearFormatting();
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        font.clearFormatting();
        font.setSize(12);

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        for (int i = 0; i <MAPKEY.size(); i++) {
            if (i!=0&&i%2==0){
                builder.endRow();
            }
            String key = MAPKEY.get(i);

            builder.insertCell();
            builder.write(TITLE.get(key));
            builder.insertCell();
            builder.write(contents.get(key).toString());
            if ("projectName".equals(key)){
                builder.getCellFormat().setHorizontalMerge(CellMerge.FIRST);
                builder.insertCell();
                builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
                builder.insertCell();
                builder.getCellFormat().setHorizontalMerge(CellMerge.PREVIOUS);
            }
        }

        builder.endTable();
    }

    private static void subjectTitle(JSONObject contents ,DocumentBuilder builder){
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        paragraphFormat.clearFormatting();

        paragraphFormat.setFirstLineIndent(32);

        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        Font font = builder.getFont();
        font.setSize(16);
        font.setName("仿宋_GB2312");
        font.setColor(Color.BLACK);
        font.setBold(true);

        builder.writeln(contents.get("questionNumber")+"."+contents.get("problem"));
        font.clearFormatting();

    }
    private static void multipleChoiceOption(JSONArray contents ,DocumentBuilder builder) throws Exception {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        paragraphFormat.clearFormatting();

        paragraphFormat.setFirstLineIndent(32);

        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        Font font = builder.getFont();
        font.setSize(16);
        font.setName("仿宋_GB2312");
        font.setColor(Color.BLACK);
        for (int i = 0; i < contents.size(); i++) {
            JSONObject content = JSONUtil.parseObj(contents.get(i));
            if ("true".equals(content.get("choose"))){
                builder.moveToMergeField("c1");
                builder.getFont().setName("Wingdings 2");
                builder.write("\u0098");
            }else {
                builder.moveToMergeField("c1");
                builder.getFont().setName("Wingdings 2");
                builder.write("\u0099");
            }
            builder.writeln(content.get("questionNumber")+"."+content.get("problem"));

        }
        contents.forEach(info ->{

        });
        font.clearFormatting();

    }
    public static String data(){
        return "{\n" +
                "\t\"surveyTime\": \"2023-8-1\",\n" +
                "\t\"feedbackTime\": \"2023-08-23 18:53:41\",\n" +
                "\t\"respondentType\": \"监理\",\n" +
                "\t\"targetCompany\": \"新源检查公司\",\n" +
                "\t\"baseName\": \"长兴分公司\",\n" +
                "\t\"projectState\": \"在建\",\n" +
                "\t\"projectName\": \"港口吊桥\",\n" +
                "\t\"subject\": [{\n" +
                "\t\t\"questionNumber\": \"1\",\n" +
                "\t\t\"problem\": \"这道题应该选C\",\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\t\"option\": \"我是A\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是B\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是C\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是D\",\n" +
                "\t\t\t\t\"choose\": true\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"checkBox\": [{\n" +
                "\t\t\t\t\"option\": \"我不满意\",\n" +
                "\t\t\t\t\"choose\": true\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我也不满意\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我特别不满意\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我~ +1\",\n" +
                "\t\t\t\t\"choose\": true\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"rest\": \"我是一个暴躁两年半的客户，我对这个玩意特别不满意，要改！！！全都要改！！！\"\n" +
                "\n" +
                "\t}, {\n" +
                "\t\t\"questionNumber\": \"2\",\n" +
                "\t\t\"problem\": \"这道题应该选B\",\n" +
                "\t\t\"multipleChoice\": [{\n" +
                "\t\t\t\t\"option\": \"我是A\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是B\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是C\",\n" +
                "\t\t\t\t\"choose\": true\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"option\": \"我是D\",\n" +
                "\t\t\t\t\"choose\": false\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"checkBox\": [],\n" +
                "\t\t\"rest\": \"\"\n" +
                "\t}]\n" +
                "}";
    }
}
