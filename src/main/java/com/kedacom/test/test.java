package com.kedacom.test;

import com.kedacom.util.CustomXWPFDocument;
import com.kedacom.util.DocUtilv2;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 15:50
 * @apiNote : TODO
 */
public class test {
    public static void main(String[] args) {
        String path = "C:\\Users\\lzp\\Desktop\\doc测试\\纯表格测试.docx";
        Map<String, Object> param = new HashMap<>();
        List<List<Object>> listList = new ArrayList<>();
        List<Object> stringList = new ArrayList<>();
        stringList.add("1");
        stringList.add("2");
        stringList.add("3");
        stringList.add("4");
        stringList.add("5");
        listList.add(stringList);
        List<Object> stringList1 = new ArrayList<>();
        stringList1.add("第二行第一个");
        listList.add(stringList1);
        param.put("JSZJQD", listList);
        String s = "C:\\Users\\lzp\\Desktop\\doc测试\\纯表格测试" + (new Date()).getTime() + ".docx";
        File file = new File(s);
        try (FileOutputStream fos = new FileOutputStream(file);
             CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(path))) {
            List<XWPFTable> tables = document.getTables();
            XWPFTable xwpfTable = tables.get(0);
            List<XWPFTableRow> rows = xwpfTable.getRows();
            if (listList.size() > 1 && rows.size() >= 1) {
                insertRowAndCopyStyle(rows.get(0), 1, 1,listList, tables.get(0));
            }
            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow xwpfTableRow = rows.get(i);
                List<XWPFTableCell> tableCells = xwpfTableRow.getTableCells();
                for (int j = 0; j < tableCells.size(); j++) {

                }
            }
            //if (tables.size()==0){
            //    XWPFTable table = document.createTable();
            //    table.setWidth(8973);
            //    //table.setCellMargins(0,0,0,0);
            //    //table.setStyleID();
            //    //table.setInsideVBorder(, , , );
            //
            //    //表格居中
            //    CTJc jc = table.getCTTbl().getTblPr().getJc();
            //    if (jc==null){
            //        jc = table.getCTTbl().getTblPr().addNewJc();
            //    }
            //    jc.setVal(STJc.CENTER);
            //    //刚刚创建的表格类默认会有一行row,row内默认会有一个单元格cell
            //    List<XWPFTableRow> rows = table.getRows();
            //    if (rows.size()!=0){
            //        XWPFTableRow xwpfTableRow = rows.get(0);
            //        xwpfTableRow.setHeight(900);
            //        for (int i = 0; i < 5; i++) {
            //            XWPFTableCell cell = xwpfTableRow.getCell(i);
            //            if (cell==null){
            //                 cell = xwpfTableRow.createCell();
            //            }
            //            cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            //            //cell可以直接设置文本,cell默认带一个<w:p>标签;在run里创建相当于会再创建一个<w:p>标签,相当于脱裤子放屁了
            //            //XWPFParagraph xwpfParagraph = cell.addParagraph();
            //            //XWPFRun run = xwpfParagraph.createRun();
            //            //run.setText("文本"+i+"runs");
            //            cell.setText("文本"+i+"cell");
            //        }
            //    }
            //}
            document.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /** 插入行,复制样式
     * @param sourceRow 复制样式的行,如果为空则为默认格式
     * @param loopCount 循环创建的次数,如果为空则默认为1
     * @param insertTablePos 插入表格的位置,如果为空则从表格最后一行添加
     * @param table 需要插入的表格,如果为空则抛出异常
     */
    private static void insertRowAndCopyStyle(XWPFTableRow sourceRow,
                                              int loopCount,
                                              int insertTablePos,
                                              List<List<Object>> insertData,
                                              XWPFTable table) {
        //todo 需要参数,源格式行,填充开始行,结束行(表格大小,数据填充范围),数据,表格本身
        //行样式
        CTTrPr rowPr = sourceRow.getCtRow().getTrPr();
        //单元格样式
        List<CTTcPr> cellCprList =new ArrayList<>();
        //段落样式
        List<CTPPr> phPprList=new ArrayList<>();
        //字体
        List<String> runsFontFamily=new ArrayList<>();
        CTTcPr cellPr = null;

        //字体大小
        Integer fontSize = null;

        //获取格式,每个单元格的格式和单元格对应的第一个段落的格式,以及第一个run的字体
        List<XWPFTableCell> tableCells = sourceRow.getTableCells();
        for (XWPFTableCell tableCell : tableCells) {
           cellCprList.add(tableCell.getCTTc().getTcPr());
            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
            if (paragraphs.size()>0){
                phPprList.add(paragraphs.get(0).getCTP().getPPr());
                List<XWPFRun> xwpfRuns = paragraphs.get(0).getRuns();
                if (xwpfRuns.size()>0) {
                    runsFontFamily.add(xwpfRuns.get(0).getFontFamily());
                }
            }
        }
        //开始复制
        //判断数据量是否大于表格内容,如果大于表格则需要额外创建空行
        for (int i = insertData.size() - 1; i >= 0; i--) {
        }
    }
}

class test2 {
    public static void main(String[] args) {
        String path = "C:\\Users\\lzp\\Desktop\\doc测试\\多个复选框同一行.docx";
        Map<String, Object> param = new HashMap<>();
        param.put("JDLX", "--成功替换!--");
        DocUtilv2.replaceWordCode(param, path);
    }
}

class test3 {
    public static void main(String[] args) {

        String str = "这是鉴定用途:${JYT}  这是第二次:${JDT} 这是另一个${JD}  第三次残缺的${";
        String reg = "\\$\\{[A-Z]+}";
        Pattern compile = Pattern.compile(reg);
        Matcher matcher = compile.matcher(str);
        while (matcher.find()) {
            int start = matcher.start();
            String group = matcher.group();
            int end = matcher.end();
            System.out.println(start + "---" + group + "---" + end);
        }
    }
}

class testLinkHashMap {
    public static void main(String[] args) {
        LinkedHashMap<Integer, String> map = new LinkedHashMap<>();
        map.put(1, "文本1");
        map.put(2, "文本2");

        map.put(3, "文本3");
        map.put(4, "文本4");
        map.put(5, "文本5");
        map.forEach((integer, s) -> System.out.println("key:" + integer + ",value:" + s));
        map.remove(3);
        map.forEach((integer, s) -> System.out.println("key:" + integer + ",value:" + s));
    }
}

class testLinkArrayList {
    public static void main(String[] args) {
        List<String> runs = new ArrayList<>();
        String con = "FLAG_TRUEFLAG_TRUE";
        runs.add("这是文本");
        runs.add(con);
        runs.add("这是二段文本" + con + "这是二段文本后续");
        for (int i = 0; i < runs.size(); i++) {
            String nowRuns = runs.get(i);
            String reg = "FLAG_TRUE";
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(nowRuns);
            if (matcher.find()) {
                String substring = nowRuns.substring(matcher.start(), matcher.end());
                if (substring.length() == nowRuns.length()) {
                    runs.set(i, "{替换成功}");
                } else {
                    String beforeText = nowRuns.substring(0, matcher.start());
                    String afterText = nowRuns.substring(matcher.end());
                    if (!StringUtils.isEmpty(afterText)) {
                        runs.set(i, afterText);
                    }
                    runs.add(i, "这是新增的勾选框");
                    if (!StringUtils.isEmpty(beforeText)) {
                        runs.add(i, beforeText);
                    }
                }
            }
        }
        runs.forEach(System.out::println);
    }
}