package com.rw.traning;

import com.alibaba.fastjson.JSON;
import com.aspose.words.*;
import com.aspose.words.Font;
import com.rw.utile.ExcelUtils;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

public class AsposeWordsUtile {


    private static final List<String> BASESSATISFICING = new ArrayList<String>(){{
        this.add("baseName");
        this.add("satisfying");
    }};

    private static final List<String>  CUSTOMERINPUT = new ArrayList<String>(){{
       this.add("projectName");
       this.add("target");
       this.add("cause");
       this.add("feedbackNumber");
    }};


    private ExcelUtils excelUtils;

    @Test
    public void trening() throws Exception {

        if (!excelUtils.GetWordLicense()){
            return;
        }


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

        String dataTotal = dataTotal();
        Map dataTotalMap = JSON.parseObject(dataTotal, Map.class);

//        List<HashMap<String, Object>> hashMaps = dataCreate1();
        //总标题
        totalTitle("上海振华重工"+dataTotalMap.get("year")+"年度产品"+dataTotalMap.get("projectStatus")+"阶段",builder);
        totalTitle("顾客满意度报告",builder);
        builder.writeln();
        //一级标题
        firstLevelTitle("一、顾客满意管理基本情况",builder);
        //二级标题
        secondLevelTitle("（一）总体情况",builder);
        //正文
        textPart(dataTotalMap.get("year")+"年，公司"+dataTotalMap.get("baseNumber")+"家单位开展了产品"+
                dataTotalMap.get("projectStatus")+"阶段的顾客满意度调查，发出顾客满意度调查表"+dataTotalMap.get("sendNumber")+
                "份，回收调查表"+dataTotalMap.get("recycleNumber")+"份，涉及调查单位"+dataTotalMap.get("supervisorNumber")+
                "家，调查覆盖率"+dataTotalMap.get("coverage")+"%。根据各单位的顾客满意度数值进行加权平均，得出公司产品"+dataTotalMap.get("projectStatus")+
                "阶段平均顾客满意度为"+dataTotalMap.get("CSI")+"%。",builder);
        //表格控件
        Table table = builder.startTable();
        //总司数据
        List<Map> basesCSI = JSON.parseArray(dataTotalMap.get("basesCSI").toString(), Map.class);
        List<String> basesNamePreprocess = basesCSI.stream().map(baseCSI -> baseCSI.get("baseName").toString()).collect(Collectors.toList());

        List<Double> satisfyingPreprocess = basesCSI.stream().map(baseCSI -> Double.parseDouble(baseCSI.get("satisfying").toString())).collect(Collectors.toList());

        String[] basesName = basesNamePreprocess.toArray(new String[0]);
        double[] satisfying = satisfyingPreprocess.stream().mapToDouble(Double::new).toArray();
        basesCSI.stream().forEach(info -> {
            info.put("satisfying",info.remove("satisfying")+"%");
        });

        String[] baseTitle = {"单位","顾客满意度"};
//        标题生成
        titleTable(baseTitle,builder,table);

//        内容生成
        contentTable(basesCSI,BASESSATISFICING,builder,table);

        builder.writeln();
        //图标生成
        graph("各单位满意度",basesName,satisfying,builder);

        List<Map> basesCase = JSON.parseArray(dataTotalMap.get("basesCase").toString(), Map.class);

        builder.writeln();
        secondLevelTitle("（二）各单位情况",builder);
        for (int i = 0; i < basesCase.size(); i++) {

            threeLevelTitle((i+1)+"."+basesCase.get(i).get("baseName").toString(),builder);
            textPart(dataTotalMap.get("year")+"年，"+basesCase.get(i).get("sendNumber")+"单位开展了产品"+dataTotalMap.get("projectStatus")+"阶段的顾客满意度调查，" +
                    "发出顾客满意度调查表"+basesCase.get(i).get("sendNumber")+"份，" +
                    "回收调查表"+basesCase.get(i).get("recycleNumber")+"份，涉及调查单位"+basesCase.get(i).get("supervisorNumber")+"家，" +
                    "调查覆盖率"+basesCase.get(i).get("coverage")+"%。",builder);

            Object obj = basesCase.get(i).get("targets");
            String string = JSON.toJSON(obj.toString()).toString();
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

        List<Map> opinionTotal = JSON.parseArray(dataTotalMap.get("opinionTotal").toString(), Map.class);
        secondLevelTitle("1.总体情况",builder);
        String[] titleOpinion = {"项目名称", "指标", "不满意原因", "反馈次数"};
        titleTable(titleOpinion,builder,table);
        contentTable(opinionTotal,CUSTOMERINPUT,builder,table);
        builder.writeln();

        Map map = JSON.parseObject(dataTotalMap.get("basesOpinion").toString(), Map.class);
        List<String> baseOpinionTitle = new ArrayList<String>(map.keySet());
        for (int i = 0; i <baseOpinionTitle.size(); i++) {
            secondLevelTitle((i+2)+"."+baseOpinionTitle.get(i),builder);
            titleTable(titleOpinion,builder,table);
            List<Map> baseContentList = JSON.parseObject(map.get(baseOpinionTitle.get(i)).toString(),List.class);

            contentTable(baseContentList,CUSTOMERINPUT,builder,table);
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
        builder.writeln();
        builder.writeln();
        builder.writeln();
        builder.writeln();
        builder.writeln();

        nodes.save("E:\\office\\"+dataTotalMap.get("year")+"产品"+dataTotalMap.get("projectStatus")+"阶段顾客满意度报告.docx");
        Date date1 = new Date();
        System.out.println(date1.getTime()-date.getTime());
    }

    /**
     * 总标题
     * @param builder
     */
    public void totalTitle(String totalTitle,DocumentBuilder builder){
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
        builder.writeln(totalTitle);
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

        paragraphFormat.setAddSpaceBetweenFarEastAndAlpha(true);
        paragraphFormat.setAddSpaceBetweenFarEastAndDigit(true);
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
    private void contentTable(List<Map> contents,List<String> contentsTitle,DocumentBuilder builder,Table table) throws Exception {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        Font font = builder.getFont();

        paragraphFormat.clearFormatting();
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        font.clearFormatting();
        font.setSize(12);

        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        contents.forEach(content -> {
            contentsTitle.forEach( title ->{
                builder.insertCell();
                builder.write(content.get(title).toString());
            });
            builder.endRow();
        });


        builder.endTable();
    }


    public String dataTotal(){

        return "{\n" +
                "\t\"year\": 2023,\n" +
                "\t\"baseNumber\": 6,\n" +
                "\t\"projectStatus\": \"在建\",\n" +
                "\t\"sendNumber\": 20,\n" +
                "\t\"recycleNumber\": 17,\n" +
                "\t\"supervisorNumber\": 5,\n" +
                "\t\"coverage\": 94,\n" +
                "\t\"CSI\": 94,\n" +
                "\t\"basesCSI\": [{\n" +
                "\t\t\"baseName\": \"<基地名1>\",\n" +
                "\t\t\"satisfying\": 74\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名2>\",\n" +
                "\t\t\"satisfying\": 97.5\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名3>\",\n" +
                "\t\t\"satisfying\": 87.2\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名4>\",\n" +
                "\t\t\"satisfying\": 78.5\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名5>\",\n" +
                "\t\t\"satisfying\": 24.5\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名6>\",\n" +
                "\t\t\"satisfying\": 87.5\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名7>\",\n" +
                "\t\t\"satisfying\": 99.1\n" +
                "\t}],\n" +
                "\t\"basesCase\": [{\n" +
                "\t\t\"baseName\": \"<基地名1>\",\n" +
                "\t\t\"sendNumber\": 7,\n" +
                "\t\t\"recycleNumber\": 6,\n" +
                "\t\t\"supervisorNumber\": 5,\n" +
                "\t\t\"coverage\": 45.5,\n" +
                "\t\t\"targets\": {\n" +
                "\t\t\t\"<指标名1>\": 98.4,\n" +
                "\t\t\t\"<指标名2>\": 45.7,\n" +
                "\t\t\t\"<指标名3>\": 85.1,\n" +
                "\t\t\t\"<指标名4>\": 78.5,\n" +
                "\t\t\t\"<指标名5>\": 75.8,\n" +
                "\t\t\t\"<指标名6>\": 87.8,\n" +
                "\t\t\t\"<指标名7>\": 97.5\n" +
                "\t\t}\n" +
                "\t}, {\n" +
                "\t\t\"baseName\": \"<基地名2>\",\n" +
                "\t\t\"sendNumber\": 20,\n" +
                "\t\t\"recycleNumber\": 15,\n" +
                "\t\t\"supervisorNumber\": 7,\n" +
                "\t\t\"coverage\": 67.2,\n" +
                "\t\t\"targets\": {\n" +
                "\t\t\t\"<指标名1>\": 98.4,\n" +
                "\t\t\t\"<指标名2>\": 45.7,\n" +
                "\t\t\t\"<指标名3>\": 85.1,\n" +
                "\t\t\t\"<指标名4>\": 78.5,\n" +
                "\t\t\t\"<指标名5>\": 75.8,\n" +
                "\t\t\t\"<指标名6>\": 87.8,\n" +
                "\t\t\t\"<指标名7>\": 97.5\n" +
                "\t\t}\n" +
                "\t}],\n" +
                "\t\"opinionTotal\": [{\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}, {\n" +
                "\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t}],\n" +
                "\t\"basesOpinion\": {\n" +
                "\t\t\"<基地名1>\": [{\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}],\n" +
                "\t\t\"<基地名2>\": [{\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}],\n" +
                "\t\t\"<基地名3>\": [{\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}],\n" +
                "\t\t\"<基地名4>\": [{\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}],\n" +
                "\t\t\"<基地名5>\": [{\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}, {\n" +
                "\t\t\t\"projectName\": \"<项目名称1>\",\n" +
                "\t\t\t\"target\": \"<指标1>\",\n" +
                "\t\t\t\"cause\": \"<满意原因1>\",\n" +
                "\t\t\t\"feedbackNumber\": \"<反馈次数>\"\n" +
                "\t\t}]\n" +
                "\t}\n" +
                "}";
    }
}
