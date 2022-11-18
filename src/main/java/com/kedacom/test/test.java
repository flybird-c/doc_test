package com.kedacom.test;

import com.kedacom.constant.KeyConstant;
import com.kedacom.util.CustomXWPFDocument;
import com.kedacom.util.DocUtilv2;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
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
            insertRowAndCopyStyle(rows.get(0), 1, 2, listList, tables.get(0), null, null);
            document.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 循环插入表格数据,行不够的时候会创建行,样式会复制sourceRow的样式
     *
     * @param sourceRow     复制样式的行,如果为空则为默认格式
     * @param startRowIndex 循环创建的次数,如果为空则默认为1
     * @param endRowIndex   插入表格的位置,如果为空则从表格最后一行添加
     * @param fontFamily    字体
     * @param tableList     要插入的表格数据
     * @param fontSize      字体大小
     * @param table         需要插入的表格,如果为空则抛出异常
     */
    private static void insertRowAndCopyStyle(XWPFTableRow sourceRow,
                                              int startRowIndex,
                                              int endRowIndex,
                                              List<List<Object>> tableList,
                                              XWPFTable table,
                                              String fontFamily,
                                              Integer fontSize) {
        //行样式
        CTTrPr rowPr = sourceRow.getCtRow().getTrPr();
        //段落样式
        CTPPr phPpr = null;
        //单元格样式
        List<CTTcPr> cellCprList = new ArrayList<>();
        //字体
        List<String> runsFontFamily = new ArrayList<>();
        //字体大小
        List<Integer> runsFontSize = new ArrayList<>();
        //获取格式
        List<XWPFTableCell> sourceRowTableCells = sourceRow.getTableCells();
        for (XWPFTableCell tableCell : sourceRowTableCells) {
            cellCprList.add(tableCell.getCTTc().getTcPr());
            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
            if (paragraphs.size() > 0) {
                phPpr = paragraphs.get(0).getCTP().getPPr();
                List<XWPFRun> xwpfRuns = paragraphs.get(0).getRuns();
                if (xwpfRuns.size() > 0) {
                    runsFontFamily.add(xwpfRuns.get(0).getFontFamily());
                    runsFontSize.add(xwpfRuns.get(0).getFontSize());
                }
            }
        }
        //判断数据量是否大于表格内容,如果大于表格则需要额外创建空行
        int tableRowSize = endRowIndex - startRowIndex;
        for (int rowIndex = 0; rowIndex < tableList.size(); rowIndex++) {
            if (rowIndex >= tableRowSize) {
                XWPFTableRow xwpfTableRow = table.insertNewTableRow(rowIndex + startRowIndex);
                xwpfTableRow.getCtRow().setTrPr(rowPr);
                for (int j = 0; j < sourceRowTableCells.size(); j++) {
                    Integer rowFontSize = fontSize;
                    String rowFontFamily = fontFamily;
                    CTTcPr ctTcPr = null;
                    if (fontSize == null && runsFontSize.size() > j) {
                        rowFontSize = runsFontSize.get(j);
                    }
                    if (fontFamily == null && runsFontFamily.size() > j) {
                        rowFontFamily = runsFontFamily.get(j);
                    }
                    if (cellCprList.size() > j) {
                        ctTcPr = cellCprList.get(j);
                    }
                    createNewCell(xwpfTableRow, ctTcPr, phPpr, rowFontFamily, rowFontSize);
                }
            }
            XWPFTableRow row = table.getRow(rowIndex + startRowIndex);
            List<XWPFTableCell> targetCell = row.getTableCells();
            for (int targetCellIndex = 0; targetCellIndex < targetCell.size(); targetCellIndex++) {
                XWPFTableCell cell = targetCell.get(targetCellIndex);
                setCellPhRunsStyleAndText(fontFamily, fontSize, runsFontFamily, runsFontSize, tableList.get(rowIndex), targetCellIndex, cell);
            }
        }
    }

    private static void setCellPhRunsStyleAndText(String fontFamily,
                                                  Integer fontSize,
                                                  List<String> runsFontFamily,
                                                  List<Integer> runsFontSize,
                                                  List<Object> rowData,
                                                  int targetCellIndex,
                                                  XWPFTableCell cell) {
        XWPFParagraph paragraph;
        if (cell.getParagraphs().size() > 0) {
            paragraph = cell.getParagraphs().get(0);
        } else {
            paragraph = cell.addParagraph();
        }

        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun targetRun;
        if (runs.size() == 0) {
            targetRun = paragraph.createRun();
        } else {
            targetRun = runs.get(0);
        }

        if (rowData.size() > targetCellIndex) {
            targetRun.setText(rowData.get(targetCellIndex) != null ? rowData.get(targetCellIndex).toString() : "");
        } else {
            targetRun.setText("");
        }
        //设置字体大小,传参优先
        if (fontSize != null) {
            targetRun.setFontSize(fontSize);
        } else if (runsFontSize.size() > targetCellIndex && runsFontSize.get(targetCellIndex) != null) {
            targetRun.setFontSize(runsFontSize.get(targetCellIndex));
        }
        //设置字体,传参优先
        if (fontFamily != null) {
            targetRun.setFontFamily(fontFamily);
        } else if (runsFontFamily.size() > targetCellIndex && runsFontFamily.get(targetCellIndex) != null) {
            targetRun.setFontFamily(runsFontFamily.get(targetCellIndex));
        }
    }

    private static void createNewCell(XWPFTableRow xwpfTableRow, CTTcPr cellPr, CTPPr phPr, String fontFamily, Integer fontSize) {
        XWPFTableCell cell = xwpfTableRow.createCell();
        cell.getCTTc().setTcPr(cellPr);

        XWPFParagraph xwpfParagraph = cell.getParagraphs().get(0);
        xwpfParagraph.getCTP().setPPr(phPr);
        XWPFRun run = xwpfParagraph.createRun();
        run.setFontFamily(fontFamily);
        run.setFontSize(fontSize);
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
class TestConst{
    @SneakyThrows
    public static void main(String[] args) {
        Field[] declaredFields = KeyConstant.class.getDeclaredFields();
        for (Field declaredField : declaredFields) {
            if (Modifier.isStatic(declaredField.getModifiers())) {
                declaredField.setAccessible(true);
                Object o = declaredField.get(KeyConstant.class);
                System.out.println(o);

            }
        }

    }
}