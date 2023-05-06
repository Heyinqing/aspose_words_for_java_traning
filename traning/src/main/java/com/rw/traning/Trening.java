package com.rw.traning;

import com.aspose.words.*;
import com.aspose.words.Font;
import com.aspose.words.Shape;
import com.rw.utile.AsposeWordsUtiles;
import org.junit.jupiter.api.Test;

import javax.swing.*;
import java.awt.*;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Set;

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
        pageSetup.setLeftMargin(50);
        pageSetup.setRightMargin(50);

        //段落操作
        ParagraphFormat paragraphFormat = builder.getParagraphFormat();
        //对齐
        paragraphFormat.setAlignment(ParagraphAlignment.JUSTIFY);
        //行缩进
        paragraphFormat.setFirstLineIndent(8);
        //段落符号
//        paragraphFormat.setKeepTogether(true);

        //字体操作
        Font font = builder.getFont();
        font.setSize(20);
//        font.set
        font.setBold(true);
        font.setColor(Color.BLACK);
        font.setUnderline(Underline.length);

        //插入字体
        builder.write("添加文字");
        builder.writeln();
        builder.writeln("换行输入");
        builder.write("123456");

        Table table = builder.startTable();
        builder.insertCell();
        table.autoFit(AutoFitBehavior.AUTO_FIT_TO_WINDOW);
        builder.getCellFormat().setVerticalAlignment(CellVerticalAlignment.CENTER);
        builder.write("单位");
        HashMap<String, String> hashMap = new HashMap<String, String>(){{
            this.put("长兴分公司","90.10");
            this.put("南通分公司","99.10");
            this.put("振华港机重工","97.65");
            this.put("上海港机重工","99.22");
            this.put("振华重工","99.11");
            this.put("启动海洋","99.05");
            this.put("南通传动","98.61");
        }};



        Set<String> strings = hashMap.keySet();
        String[] array = strings.toArray(new String[0]);
        double[] doubles = new double[array.length];
        for (int i = 0; i < array.length; i++) {

            builder.insertCell();
            builder.write(array[i]);
        }

        builder.endRow();
        builder.insertCell();
        builder.write(" ");

        for (int i = 0; i < array.length; i++) {

            builder.insertCell();
            String text = hashMap.get(array[i]);
            doubles[i] = Double.valueOf(text);
            builder.write(text+"%");
        }
        builder.endRow();
        builder.endTable();


        Chart chart = builder.insertChart(ChartType.COLUMN, 432, 252).getChart();
        chart.getSeries().clear();

        chart.getSeries().add("Series 1",array,doubles);
        chart.getTitle().setText("Test");
        chart.getLegend().setPosition(LegendPosition.BOTTOM);

        Paragraph paragraph = new Paragraph(nodes);
        Run run = new Run(nodes);
        paragraph.appendChild(run);

        builder.write(run.getText());
        nodes.save("E:\\office\\测试文档1.docx");
    }
}
