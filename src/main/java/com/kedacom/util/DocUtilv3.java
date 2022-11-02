package com.kedacom.util;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 15:21
 * @apiNote : TODO
 */
@Slf4j
public class DocUtilv3 {
    @SneakyThrows
    public static String replaceWordCode(Map<String, Object> param, String srcWordPath) {
        checkParam(param, srcWordPath);
        String targetPath = getTargetPath(srcWordPath);
        try (FileOutputStream fos = new FileOutputStream(targetPath);
             CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(srcWordPath))) {
            //获取所有段落
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            if (!CollectionUtils.isEmpty(paragraphs)) {
                //处理item系列编码
                Map<String, Object> itemParam = handItemCodes(param);
                for (XWPFParagraph paragraph : paragraphs) {
                    paragraphHandler(paragraph, param,itemParam);
                }
            }
            //处理表格
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (XWPFTableCell tableCell : tableCells) {
                        List<XWPFParagraph> tableCellParagraphs = tableCell.getParagraphs();
                        //todo 和表格一样的处理流程,只处理文本/图片
                        for (XWPFParagraph paragraph : tableCellParagraphs) {
                            paragraphHandler(paragraph, param);
                        }
                    }
                }
            }
            //页眉
            List<XWPFHeader> headerList = document.getHeaderList();
            for (XWPFHeader xwpfHeader : headerList) {
                List<XWPFParagraph> xwpfHeaderParagraphs = xwpfHeader.getParagraphs();
                //todo 只处理文本
            }
            //页脚
            List<XWPFFooter> footerList = document.getFooterList();
            for (XWPFFooter xwpfFooter : footerList) {
                List<XWPFParagraph> xwpfFooterParagraphs = xwpfFooter.getParagraphs();
                //todo 只处理文本
            }
        }
        return null;
    }

    private static void paragraphHandler(XWPFParagraph paragraph, Map<String, Object> param, Map<String, Object> itemParam) {

    }

    private static void paragraphHandler(XWPFParagraph paragraph, Map<String, Object> param) {
        int startRunIndex=-1;
        int endRunIndex=-1;
        //获取所有runs去掉空格拼接成整段文本(一个段落)
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun xwpfRun = runs.get(i);
            String text = xwpfRun.text();
            String reg="\\$\\{[A-Z]+}";
            Pattern compile = Pattern.compile(reg);
            for (Map.Entry<String, Object> entry : param.entrySet()) {
                Matcher matcher = compile.matcher(entry.getKey());
                if (matcher.find()){
                    int end = matcher.end();
                    text= text.replace(entry.getKey(),entry.getValue().toString());
                }
            }
        }
    }

    private static String strHandler(String paragraphText, Map<String, Object> param) {
        String text=textHandler(paragraphText,param);
        text= imgHandler(text,param);
        text= choolseHandler(text,param);
        return text;
    }

    private static String choolseHandler(String text, Map<String, Object> param) {
        return null;
    }

    private static String imgHandler(String text, Map<String, Object> param) {
        return null;
    }

    /** 普通文本编码处理
     * @param paragraphText
     * @param param
     * @return
     */
    private static String textHandler(String paragraphText, Map<String, Object> param) {
        log.info("文本编码处理开始:{}",paragraphText);
        String reg="\\$\\{[A-Z]+}";
        Pattern compile = Pattern.compile(reg);
        for (Map.Entry<String, Object> entry : param.entrySet()) {
            Matcher matcher = compile.matcher(entry.getKey());
            if (matcher.find()){
                paragraphText= paragraphText.replace(entry.getKey(),entry.getValue().toString());
            }
        }
        log.info("文本编码处理结束:{}",paragraphText);
        return paragraphText;
    }

    public static Map<String, Object> handItemCodes(Map<String, Object> param) {
        Object item = param.get("ITEM");
        Map<String, Object> itemMap = new HashMap<>();
        if (!ObjectUtils.isEmpty(item)) {
            for (Map.Entry<String, Object> entry : param.entrySet()) {
                itemMap.put(entry.getKey() + "-" + item, entry.getValue());
            }
        }
        return itemMap;
    }

    private static String getTargetPath(String srcWordPath) {
        String[] split = srcWordPath.split("\\.");
        String filePaths = split[0];
        String docType = split[split.length - 1];
        return filePaths + (new Date()).getTime() + "." + docType;
    }

    private static void checkParam(Map<String, Object> param, String srcWordPath) {
        if (CollectionUtils.isEmpty(param)) {
            throw new IllegalArgumentException("参数不能为空或者为null");
        }
        String[] split = srcWordPath.split("\\.");
        String docType = split[split.length - 1];
        if ("docx".equalsIgnoreCase(docType)) {
            throw new IllegalArgumentException("不是支持的docx类型");
        }
    }
}
