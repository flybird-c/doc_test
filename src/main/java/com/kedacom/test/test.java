package com.kedacom.test;

import com.kedacom.util.CustomXWPFDocument;
import com.kedacom.util.DocUtilv2;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTJc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STJc;
import org.springframework.util.StringUtils;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
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
        String path = "C:\\Users\\lzp\\Desktop\\doc测试\\空白表格.docx";
        Map<String, Object> param = new HashMap<>();
        List<List<String>> listList=new ArrayList<>();
        List<String> stringList=new ArrayList<>();
        stringList.add("1");
        stringList.add("2");
        stringList.add("3");
        stringList.add("4");
        stringList.add("5");
        listList.add(stringList);
        List<String> stringList1=new ArrayList<>();
        stringList1.add("第二行第一个");
        listList.add(stringList1);
        param.put("JSZJQD",listList);
        String s = "C:\\Users\\lzp\\Desktop\\doc测试\\空白表格" + (new Date()).getTime() + ".docx";
        File file = new File(s);
        try (FileOutputStream fos = new FileOutputStream(file);
             CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(path))) {
            List<XWPFTable> tables = document.getTables();
            if (tables.size()==0){
                XWPFTable table = document.createTable();
                table.setWidth(8973);
                //table.setCellMargins(0,0,0,0);
                //table.setStyleID();
                //table.setInsideVBorder(, , , );
                CTJc jc = table.getCTTbl().getTblPr().getJc();
                if (jc==null){
                    jc = table.getCTTbl().getTblPr().addNewJc();
                }
                jc.setVal(STJc.CENTER);
                List<XWPFTableRow> rows = table.getRows();
                if (rows.size()!=0){
                    XWPFTableRow xwpfTableRow = rows.get(0);
                    xwpfTableRow.setHeight(900);
                    for (int i = 0; i < 5; i++) {
                        XWPFTableCell cell = xwpfTableRow.getCell(i);
                        if (cell==null){
                             cell = xwpfTableRow.createCell();
                        }
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                        //cell.setText("文本"+i);
                        XWPFParagraph xwpfParagraph = cell.addParagraph();
                        XWPFRun run = xwpfParagraph.createRun();
                        run.setText("文本"+i);
                    }
                }
            }
            document.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
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
        List<String> runs=new ArrayList<>();
        String con="FLAG_TRUEFLAG_TRUE";
        runs.add("这是文本");
        runs.add(con);
        runs.add("这是二段文本"+con+"这是二段文本后续");
        for (int i = 0; i < runs.size(); i++) {
            String nowRuns = runs.get(i);
            String reg="FLAG_TRUE";
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(nowRuns);
            if (matcher.find()) {
                String substring = nowRuns.substring(matcher.start(), matcher.end());
                if (substring.length()==nowRuns.length()){
                    runs.set(i,"{替换成功}");
                }else {
                    String beforeText=nowRuns.substring(0, matcher.start());
                    String afterText=nowRuns.substring(matcher.end());
                    if (!StringUtils.isEmpty(afterText)) {
                        runs.set(i,afterText);
                    }
                    runs.add(i, "这是新增的勾选框");
                    if (!StringUtils.isEmpty(beforeText)) {
                        runs.add(i,beforeText);
                    }
                }
            }
        }
runs.forEach(System.out::println);
    }
}