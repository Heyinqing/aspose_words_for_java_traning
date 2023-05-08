package com.rw.traning;

import com.alibaba.fastjson.JSON;
import com.aspose.words.*;
import com.aspose.words.Font;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.util.*;
import java.util.List;

public class Trening{


    @Test
    public void trening() throws Exception {

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

        HashMap<String, String> hashMap = new HashMap<String, String>(){{
            this.put("长兴分公司","90.10");
            this.put("南通分公司","99.10");
            this.put("振华港机重工","97.65");
            this.put("上海港机重工","99.22");
            this.put("振华重工","99.11");
            this.put("启动海洋","99.05");
            this.put("南通传动","98.61");
        }};
        List<HashMap<String, Object>> hashMaps = dataCreate1();
        //总标题
        totalTitle(builder);
        //一级标题
        firstLevelTitle("一、顾客满意管理基本情况",builder);
        //二级标题
        secondLevelTitle("（一）总体情况",builder);
        //正文
        textPart("{年份}年，公司{多少}家单位开展了产品{项目状态}阶段的顾客满意度调查，发出顾客满意度调查表{多少}份，回收调查表{多少}份，涉及调查单位{多少}家，调查覆盖率{百分比}。\n" +
                "根据各单位的顾客满意度数值进行加权平均，得出公司产品{项目状态}阶段平均顾客满意度为{百分比}。",builder);
        //表格控件
        Table table = builder.startTable();
        //总司数据


        //表格标题
        String[] array = hashMap.keySet().toArray(new String[0]);
        //表格数据
        String[] contentTable = new String[array.length];
        //表格y轴
        double[] doubles = new double[array.length];
        //数据拆解
        for (int i = 0; i < array.length; i++) {
            doubles[i] = Double.parseDouble(hashMap.get(array[i]));
            contentTable[i] = hashMap.get(array[i])+"%";
        }

        //标题生成
//        titleTable(array,builder,table);

        //内容生成
//        List<String[]> content = new ArrayList<>();
//        content.add(contentTable);
//        contentTable(content,builder,table);

        builder.writeln();
        //图标生成
        graph("各单位满意度",array,doubles,builder);

        builder.writeln();
        secondLevelTitle("（二）各单位情况",builder);
        for (int i = 0; i < hashMaps.size(); i++) {

            threeLevelTitle((i+1)+"."+hashMaps.get(i).get("baseName").toString(),builder);
            textPart("{年份}年，"+hashMaps.get(i).get("sendNumber")+"单位开展了产品{项目状态}阶段的顾客满意度调查，" +
                    "发出顾客满意度调查表"+hashMaps.get(i).get("sendNumber")+"份，" +
                    "回收调查表"+hashMaps.get(i).get("recycleNumber")+"份，涉及调查单位"+hashMaps.get(i).get("supervisorNumber")+"家，" +
                    "调查覆盖率"+hashMaps.get(i).get("coverage")+"。",builder);

            Object obj = hashMaps.get(i).get("targets");
            String string = JSON.toJSON(obj).toString();
            Map beanMap = JSON.parseObject(string, Map.class);
            Object[] objs = beanMap.keySet().toArray();
            String[] title = new String[objs.length];
            //表格y轴
            double[] contentY = new double[objs.length];
            //数据拆解
            for (int j = 0; j < objs.length; j++) {
              title[j] = objs[j].toString();
              contentY[j] = Double.parseDouble(beanMap.get(title[j]).toString());
            }
            graph("各项指标客户满意度",title,contentY,builder);
        };



        builder.writeln();
        builder.writeln();
        //一级标题
        firstLevelTitle("二、顾客提出的意见",builder);
        String[] titleOpinion = {"项目名称", "指标", "不满意原因", "反馈次数"};
        titleTable(titleOpinion,builder,table);
        List<Map<String, String>> maps = dataCreate2();
        List<String> contentsTitle = new ArrayList<>(maps.get(0).keySet());
        contentTable(maps,contentsTitle,builder,table);
        builder.writeln();

        Map map = dataCreate3();
        List<String> baseTitle = new ArrayList<String>(map.keySet());
        for (int i = 0; i <baseTitle.size(); i++) {
            secondLevelTitle((i+1)+"."+baseTitle.get(i),builder);
            titleTable(titleOpinion,builder,table);
            List<Map<String, String>> baseContentList = JSON.parseObject(map.get(baseTitle.get(i)).toString(),List.class);
            List<String> baseContentTitle = new ArrayList<>(baseContentList.get(0).keySet());
            contentTable(baseContentList,baseContentTitle,builder,table);
            builder.writeln();
        }
        builder.writeln();
        firstLevelTitle("三、顾客满意度分析",builder);
        builder.writeln();
        builder.writeln();
        builder.writeln();
        builder.writeln();
        builder.writeln();

        firstLevelTitle("四、改进措施和建议",builder);

//        table.setAlignment(ParagraphAlignment.CENTER);
//        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);


        Date date = new Date();
        nodes.save("E:\\office\\测试文档1.docx");
        Date date1 = new Date();
        System.out.println(date1.getTime()-date.getTime());
    }

    /**
     * 总标题
     * @param builder
     */
    public void totalTitle(DocumentBuilder builder){
        Font font = builder.getFont();
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        font.setSize(22);
        font.setBold(true);
        font.setName("黑体");
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);

        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);

        //插入字体
        builder.writeln();
        builder.writeln("上海振华重工{年份}年度产品{项目状态}阶段");
        builder.writeln("顾客满意度报告");
        builder.writeln();
        //重置字体样式
        font.clearFormatting();

    }

    /**
     * 一级标题
     * @param content
     * @param builder
     */
    private void firstLevelTitle(String content,DocumentBuilder builder){
        //字体操作
        Font font = builder.getFont();

        //段落操作
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //行缩进
        paragraphFormat.setFirstLineIndent(32);
        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);
        //段落符号
        paragraphFormat.setKeepTogether(true);

        font.setSize(16);
        font.setName("黑体");
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);

        //插入字体
        builder.writeln(content);

        //重置字体样式
        font.clearFormatting();
    }

    /**
     * 二级标题
     * @param content
     * @param builder
     */
    private void secondLevelTitle(String content,DocumentBuilder builder){
        //字体操作
        Font font = builder.getFont();
        //段落操作
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //行缩进
        paragraphFormat.setFirstLineIndent(32);
        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);
        //段落符号
        paragraphFormat.setKeepTogether(true);

        font.setSize(16);
        font.setName("楷体");
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);

        //插入字体
        builder.writeln(content);

        //重置字体样式
        font.clearFormatting();
    }

    /**
     * 三级标题
     * @param content
     * @param builder
     */
    private void threeLevelTitle(String content,DocumentBuilder builder){
        //字体操作
        Font font = builder.getFont();
        //段落操作
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //行缩进
        paragraphFormat.setFirstLineIndent(32);
        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);
        //段落符号
        paragraphFormat.setKeepTogether(true);

        font.setSize(16);
        font.setName("仿宋_GB2312");
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);
        font.setBold(true);
        //插入字体
        builder.writeln(content);
        //重置字体样式
        font.clearFormatting();
    }


    /**
     * 正文
     * @param content
     * @param builder
     */
    private void textPart(String content,DocumentBuilder builder){
        //字体操作
        Font font = builder.getFont();
        //段落操作
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        //行缩进
        paragraphFormat.setFirstLineIndent(32);
        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.length);
        //段落符号
        paragraphFormat.setKeepTogether(true);

        font.setSize(16);
        font.setName("仿宋");
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);

        //插入字体
        builder.writeln(content);
        //重置字体样式
        font.clearFormatting();
    }

    /**
     * 图表生成
     * @param content 标题
     * @param array
     * @param doubles
     * @param builder
     * @throws Exception
     */
    private void graph(String content,String[] array,double[] doubles,DocumentBuilder builder) throws Exception {

        Chart chart = builder.insertChart(ChartType.COLUMN, 432, 252).getChart();
        chart.getSeries().clear();

        chart.getSeries().add(content,array,doubles);
        chart.getTitle().setText(content);
        chart.getLegend().setPosition(LegendPosition.BOTTOM);

        builder.writeln();
    }



    /**
     * 表格标题生成
     * @param array
     * @param builder
     * @param table
     * @throws Exception
     */
    private void titleTable(String[] array,DocumentBuilder builder,Table table) throws Exception {

        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        Font font = builder.getFont();



        paragraphFormat.setAlignment(ParagraphAlignment.CENTER);
        paragraphFormat.clearFormatting();

        font.clearFormatting();
        font.setSize(12);

        builder.insertCell();
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        for (String s : array) {
            builder.insertCell();
            builder.write(s);
        }

        builder.endRow();

    }

    /**
     * 表格内容生成
     * @param contents
     * @param builder
     * @param table
     * @throws Exception
     */
    private void contentTable(List<Map<String,String>> contents,List<String> contentsTitle,DocumentBuilder builder,Table table) throws Exception {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        Font font = builder.getFont();

        paragraphFormat.clearFormatting();
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        font.clearFormatting();
        font.setSize(12);

        builder.insertCell();
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        contents.forEach(content -> {
            contentsTitle.forEach( title ->{
                builder.insertCell();
                builder.write(content.get(title));
            });
            builder.endRow();
        });


//        builder.endRow();
        builder.endTable();
    }

    private List<HashMap<String,Object>> dataCreate1(){
        return new ArrayList<HashMap<String,Object>>(){{
            this.add(new HashMap<String, Object>(){{
                this.put("baseName","长兴分公司");
                this.put("sendNumber",25D);
                this.put("recycleNumber",20D);
                this.put("supervisorNumber",10D);
                this.put("coverage",83.3);
                this.put("targets",new HashMap<String, Double>(){{
                    this.put("营销",91.2);
                    this.put("设计",99.3);
                    this.put("工艺",97.6);
                    this.put("物资",98.4);
                    this.put("生产",99.1);
                    this.put("质量",98.7);
                    this.put("安全",98.2);
                }});
            }});

            this.add(new HashMap<String, Object>(){{
                this.put("baseName","南通分公司");
                this.put("sendNumber",25D);
                this.put("recycleNumber",20D);
                this.put("supervisorNumber",10D);
                this.put("coverage",83.3);
                this.put("targets",new HashMap<String, Double>(){{
                    this.put("营销",91.2);
                    this.put("设计",99.3);
                    this.put("工艺",97.6);
                    this.put("物资",98.4);
                    this.put("生产",99.1);
                    this.put("质量",98.7);
                    this.put("安全",98.2);
                }});
            }});

            this.add(new HashMap<String, Object>(){{
                this.put("baseName","振华港机重工");
                this.put("sendNumber",25D);
                this.put("recycleNumber",20D);
                this.put("supervisorNumber",10D);
                this.put("coverage",83.3);
                this.put("targets",new HashMap<String, Double>(){{
                    this.put("营销",91.2);
                    this.put("设计",99.3);
                    this.put("工艺",97.6);
                    this.put("物资",98.4);
                    this.put("生产",99.1);
                    this.put("质量",98.7);
                    this.put("安全",98.2);
                }});
            }});

            this.add(new HashMap<String, Object>(){{
                this.put("baseName","上海港机重工");
                this.put("sendNumber",25D);
                this.put("recycleNumber",20D);
                this.put("supervisorNumber",10D);
                this.put("coverage",83.3);
                this.put("targets",new HashMap<String, Double>(){{
                    this.put("营销",91.2);
                    this.put("设计",99.3);
                    this.put("工艺",97.6);
                    this.put("物资",98.4);
                    this.put("生产",99.1);
                    this.put("质量",98.7);
                    this.put("安全",98.2);
                }});
            }});

        }};
    }


    public  List<Map<String, String>> dataCreate2(){
        String a= "{\"opinionTotal\":[\n" +
                "\t\t{\"projectName\":\"<项目名称1>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称1>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称1>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称1>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称2>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称2>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称3>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称4>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t\t{\"projectName\":\"<项目名称4>\",\"target\":\"<指标1>\",\"cause\":\"<满意原因1>\",\"feedbackNumber\":\"<反馈次数>\"},\n" +
                "\t]}";
        Map map = JSON.parseObject(a, Map.class);
        List<Map<String, String>> hashMaps = new ArrayList<>();
        Object object = map.get("opinionTotal");
        if (object instanceof ArrayList<?>) {
            for (Object o : (List<?>) object) {
                hashMaps.add(Map.class.cast(o));
            }
        }
        return hashMaps;
//        return null;
    }


    public Map dataCreate3(){
        String a = "{\n" +
                "\t\"basesOpinion\": {\n" +
                "\t\t\"<基地名1>\": [{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"<基地名2>\": [{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"<基地名3>\": [{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"<基地名4>\": [{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t}\n" +
                "\t\t],\n" +
                "\t\t\"<基地名5>\": [{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t},\n" +
                "\t\t\t{\n" +
                "\t\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t\t}\n" +
                "\t\t]\n" +
                "\t}\n" +
                "}";

        Map map = JSON.parseObject(JSON.parseObject(a, Map.class).get("basesOpinion").toString(),Map.class);
        return map;
    }
}
