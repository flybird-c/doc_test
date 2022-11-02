package com.kedacom.test;

import com.kedacom.util.CustomXWPFDocument;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.File;
import java.io.FileInputStream;
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
    @SneakyThrows
    public static void main(String[] args) {
        String path="C:\\Users\\lzp\\Desktop\\doc测试\\test.docx";
        File file = new File(path);
        if (!file.exists()){
            File file1 = new File(file.getParent());
            file1.mkdirs();
            file.createNewFile();
        }
        FileInputStream inputStream = new FileInputStream(path);
        CustomXWPFDocument customXWPFDocument = new CustomXWPFDocument(inputStream);
        List<XWPFParagraph> paragraphs = customXWPFDocument.getParagraphs();
        XWPFParagraph xwpfParagraph = paragraphs.get(0);

        XWPFRun run = xwpfParagraph.createRun();
        run.setFontFamily("Wingdings 2");
        run.setText("\u0052 测试");
        xwpfParagraph.addRun(run);

    }
}
class test2{
    public static void main(String[] args) {
        String str="这是鉴定用途:${JDYT2}    第二次${JDYT2}";
        int i = str.indexOf("${", -1);

        char c = str.charAt(7);
        System.out.println(c);
        System.out.println(i);
        String x = str.replace("${JDYT2}", "行政");
        System.out.println(x);
    }
}
class test3{
    public static void main(String[] args) {

        String str="这是鉴定用途:${JDYT}  这是第二次:${JDYT} 这是另一个${JD}  第三次残缺的${";
        String reg="\\$\\{[A-Z]+}";
        Pattern compile = Pattern.compile(reg);
        Matcher matcher = compile.matcher(str);
        if (matcher.find()) {
            String group1 = matcher.group();
            System.out.println(group1);
            for (int i = 0; i < matcher.groupCount(); i++) {
                int start = matcher.start(i);
                String group = matcher.group(i);
                int end = matcher.end(i);
                System.out.println(start+"---"+group+"---"+end);
            }
        }
    }
}