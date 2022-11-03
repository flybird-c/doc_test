package com.kedacom.test;

import com.kedacom.util.CustomXWPFDocument;
import com.kedacom.util.DocUtilv2;
import lombok.SneakyThrows;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.awt.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 15:50
 * @apiNote : TODO
 */
public class test {
    public static void main(String[] args)  {
        String path="C:\\Users\\lzp\\Desktop\\doc测试\\多个复选框同一行.docx";
        Map<String, Object> param=new HashMap<>();
        param.put("${JDLX}","--成功替换!--");
        String s ="C:\\Users\\lzp\\Desktop\\doc测试\\多个复选框同一行" + (new Date()).getTime() + ".docx" ;
        File file = new File(s);
        try (FileOutputStream fos = new FileOutputStream(file);
             CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(path))) {

            List<XWPFParagraph> paragraphs = document.getParagraphs();
            for (XWPFParagraph paragraph : paragraphs) {
                List<XWPFRun> runs = paragraph.getRuns();
                Map<Integer,String> idForRuns=new HashMap<>(256);
                //文本缓存,与id对应
                for (int i = 0; i < runs.size(); i++) {
                    idForRuns.put(i,runs.get(i).toString());
                }
                for (int i = 0; i < runs.size(); i++) {
                    //获取本段run的文本
                    String text = idForRuns.get(i);
                    String newText = text;
                    String reg="\\$\\{\\S+}";
                    Pattern compile = Pattern.compile(reg);
                    for (Map.Entry<String, Object> entry : param.entrySet()) {
                        //匹配文本key
                        Matcher matcher = compile.matcher(entry.getKey());
                        while (matcher.find()){
                            newText= newText.replace(entry.getKey(),entry.getValue().toString());
                        }
                    }
                    XWPFRun xwpfRun = runs.get(i);
                    xwpfRun.setText(newText,0);
                    //paragraph.removeRun(i);
                    //XWPFRun xwpfRun1 = paragraph.insertNewRun(i);
                    //xwpfRun1.setText(newText);
                }
            }
            document.write(fos);
        }catch (Exception e){
            e.printStackTrace();
        }
    }
}
class test2{
    public static void main(String[] args) {
        String path="C:\\Users\\lzp\\Desktop\\doc测试\\多个复选框同一行.docx";
        Map<String, Object> param=new HashMap<>();
        param.put("JDLX","--成功替换!--");
        DocUtilv2.replaceWordCode(param,path);
    }
}
class test3{
    public static void main(String[] args) {

        String str="这是鉴定用途:${JYT}  这是第二次:${JDT} 这是另一个${JD}  第三次残缺的${";
        String reg="\\$\\{[A-Z]+}";
        Pattern compile = Pattern.compile(reg);
        Matcher matcher = compile.matcher(str);
        while (matcher.find()) {
            int start = matcher.start();
            String group = matcher.group();
            int end = matcher.end();
            System.out.println(start+"---"+group+"---"+end);
        }
    }
}