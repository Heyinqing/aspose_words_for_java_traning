package com.rw.traning;

import com.aspose.words.*;
import com.aspose.words.Font;
import org.junit.jupiter.api.Test;

import javax.swing.*;
import java.awt.*;
import java.util.*;
import java.util.List;

public class Trening{


    @Test
    public void trening() throws Exception {

        Document nodes = new Document();
        DocumentBuilder builder = new DocumentBuilder(nodes);
        //修改样式
        PageSetup pageSetup = builder.getPageSetup();
        pageSetup.setPaperSize(PaperSize.A4);
        pageSetup.setOrientation(Orientation.PORTRAIT);
        pageSetup.setVerticalAlignment(PageVerticalAlignment.TOP);
        pageSetup.setLeftMargin(90);
        pageSetup.setTopMargin(72);
        pageSetup.setBottomMargin(72 );
        pageSetup.setRightMargin(90);


        totalTitle(builder);

        firstLevelTitle("一、顾客满意管理基本情况",builder);

        secondLevelTitle("（一）总体情况",builder);

        textPart("{年份}年，公司{多少}家单位开展了产品{项目状态}阶段的顾客满意度调查，发出顾客满意度调查表{多少}份，回收调查表{多少}份，涉及调查单位{多少}家，调查覆盖率{百分比}。\n" +
                "根据各单位的顾客满意度数值进行加权平均，得出公司产品{项目状态}阶段平均顾客满意度为{百分比}。",builder);

        Table table = builder.startTable();

        HashMap<String, String> hashMap = new HashMap<String, String>(){{
            this.put("长兴分公司","90.10");
            this.put("南通分公司","99.10");
            this.put("振华港机重工","97.65");
            this.put("上海港机重工","99.22");
            this.put("振华重工","99.11");
            this.put("启动海洋","99.05");
            this.put("南通传动","98.61");
        }};

        //表格标题
        String[] array = hashMap.keySet().toArray(new String[0]);

        String[] contentTable = new String[array.length];
        //表格y轴
        double[] doubles = new double[array.length];

        for (int i = 0; i < array.length; i++) {
            doubles[i] = Double.parseDouble(hashMap.get(array[i]));
            contentTable[i] = hashMap.get(array[i])+"%";
        }

        //标题生成
        titleTable(array,builder,table);

        //内容生成
        List<String[]> content = new ArrayList<>();
        content.add(contentTable);
        contentTable(content,builder,table);


        builder.writeln();

        graph("各单位满意度",array,doubles,builder);

        builder.writeln();
        secondLevelTitle("（二）各单位情况",builder);


        for (int i = 0; i < 3; i++) {
            threeLevelTitle((i+1)+".{单位名称}",builder);
            textPart("{年份}年，{单位名称}单位开展了产品{项目状态}阶段的顾客满意度调查，发出顾客满意度调查表{发出}份，回收调查表{回收}份，涉及调查单位{产于}家，调查覆盖率{百分比}。",builder);
        }


        nodes.save("E:\\office\\测试文档1.docx");
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
     * 图标生成
     * @param content 标题
     * @param array
     * @param doubles
     * @param builder
     * @throws Exception
     */
    public void graph(String content,String[] array,double[] doubles,DocumentBuilder builder) throws Exception {

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
    public void titleTable(String[] array,DocumentBuilder builder,Table table) throws Exception {

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
    private void contentTable(List<String[]> contents,DocumentBuilder builder,Table table) throws Exception {
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();

        Font font = builder.getFont();

        paragraphFormat.clearFormatting();
        paragraphFormat.setAlignment(ParagraphAlignment.length);

        font.clearFormatting();
        font.setSize(12);

        builder.insertCell();
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);

        contents.forEach(content -> {
            for (String aDouble : content) {
                builder.insertCell();
                builder.write(aDouble);
            }
        });


        table.setAlignment(ParagraphAlignment.CENTER);
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);

        builder.endRow();
        builder.endTable();
    }
}
