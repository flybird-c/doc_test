package com.kedacom.util;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.*;
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
            //处理item系列编码
            Map<String, Object> itemParam = handItemCodes(param);
            param.putAll(itemParam);
            //获取所有段落
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            if (!CollectionUtils.isEmpty(paragraphs)) {
                for (XWPFParagraph paragraph : paragraphs) {
                    paragraphHandler(paragraph, param);
                }
            }
            //处理表格 还有表格增行操作
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (XWPFTableCell tableCell : tableCells) {
                        List<XWPFParagraph> tableCellParagraphs = tableCell.getParagraphs();
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
                for (XWPFParagraph paragraph : xwpfHeaderParagraphs) {
                    paragraphHandler(paragraph, param);
                }
            }
            //页脚
            List<XWPFFooter> footerList = document.getFooterList();
            for (XWPFFooter xwpfFooter : footerList) {
                List<XWPFParagraph> xwpfFooterParagraphs = xwpfFooter.getParagraphs();
                for (XWPFParagraph paragraph : xwpfFooterParagraphs) {
                    paragraphHandler(paragraph, param);
                }
            }
            document.write(fos);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return targetPath;
    }

    private static void paragraphHandler(XWPFParagraph paragraph, Map<String, Object> param) {
        //获取所有runs去掉空格拼接成整段文本(一个段落)
        List<XWPFRun> runs = paragraph.getRuns();
        List<String> idForRuns = new ArrayList<>(256);
        //文本缓存,与id对应
        for (int i = 0; i < runs.size(); i++) {
            idForRuns.add(runs.get(i).toString());
        }
        //跨行文本预处理
        multilineCodeHandler(paragraph, runs, idForRuns);
        //文本编码处理
        textCodeHandler(param, runs, idForRuns);
        //todo 勾选框处理

        //todo 图片编码处理
    }

    private static void textCodeHandler(Map<String, Object> param, List<XWPFRun> runs, List<String> idForRuns) {
        for (int i = 0; i < runs.size(); i++) {
            //获取本段run的文本
            String text = idForRuns.get(i);
            String newText = text;
            String reg = "\\$\\{\\w+}";
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(text);
            //匹配文本key
            while (matcher.find()) {
                for (Map.Entry<String, Object> entry : param.entrySet()) {
                    newText = newText.replace(entry.getKey(), entry.getValue().toString());
                }
            }
            //如果文本有编码被替换,则更新run
            if (!text.equals(newText)) {
                XWPFRun xwpfRun = runs.get(i);
                xwpfRun.setText(newText, 0);
                idForRuns.remove(i);
                idForRuns.add(i, newText);
            }
        }
    }

    private static void multilineCodeHandler(XWPFParagraph paragraph, List<XWPFRun> runs, List<String> idForRuns) {
        int startRunIndex = -1;
        int endRunIndex = -1;
        for (int i = 0; i < runs.size(); i++) {
            //获取本段run的文本
            String text = idForRuns.get(i);
            //匹配残缺的编码标识
            String endRunMutilatedReg = "\\$\\{?[^}]*$";
            Pattern compile = Pattern.compile(endRunMutilatedReg);
            Matcher matcher = compile.matcher(text);
            if (matcher.find()) {
                //记录残缺编码起始位置
                startRunIndex = i;
            }
            int lastRunSubIndex = -1;
            //寻找下一段匹配的编码末尾
            if (startRunIndex != -1) {
                String endRunReg = "^\\S*}";
                Pattern pattern = Pattern.compile(endRunReg);
                Matcher matcher1 = pattern.matcher(text);
                if (matcher1.find()) {
                    //记录末尾位置与字符下标
                    endRunIndex = i;
                    lastRunSubIndex = matcher1.end();
                }
            }
            //处理跨行文本
            if (startRunIndex != -1 && endRunIndex != -1) {
                StringBuilder startRunText = new StringBuilder();
                StringBuilder endRunText = new StringBuilder();
                //处理文本
                //todo 将处理后的文本与原本的样式原样设置到run,删除中间的run,更新run头和run尾 注意:移除run方法,后面的run会补位到前面
                XWPFRun startRun = runs.get(startRunIndex);
                XWPFRun endRun = runs.get(endRunIndex);
                for (int index = startRunIndex; index <= endRunIndex; index++) {
                    //中间的run属于编码部分
                    if (index < endRunIndex) {
                        startRunText.append(idForRuns.get(index));
                    }
                    //末尾位置需要区分编码部分与正常文本
                    if (index == endRunIndex) {
                        String lastRunText = idForRuns.get(index);
                        //}之后的文本
                        String substrAfter1 = lastRunText.substring(lastRunSubIndex);
                        endRunText.append(substrAfter1);
                        //}
                        String substrBefore = lastRunText.replace(substrAfter1, "");
                        startRunText.append(substrBefore);
                    }
                    //结束时删除中间的run和缓存的文本
                    if (index > startRunIndex && index < endRunIndex) {
                        //删除后自动补位,对应下标相应-1
                        paragraph.removeRun(index);
                        endRunIndex--;
                        idForRuns.remove(index);
                        index--;
                    }
                }
                //pos代表w:t标签的下标
                startRun.setText(startRunText.toString(), 0);
                //替换内容
                idForRuns.remove(startRunIndex);
                idForRuns.add(startRunIndex, startRunText.toString());
                endRun.setText(endRunText.toString(), 0);
                idForRuns.remove(endRunIndex);
                idForRuns.add(endRunIndex, endRunText.toString());
                //重置标志位
                startRunIndex = -1;
                endRunIndex = -1;
            }
        }
    }

    private static String strHandler(String paragraphText, Map<String, Object> param) {
        String text = textHandler(paragraphText, param);
        text = imgHandler(text, param);
        text = choolseHandler(text, param);
        return text;
    }

    private static String choolseHandler(String text, Map<String, Object> param) {
        return null;
    }

    private static String imgHandler(String text, Map<String, Object> param) {
        return null;
    }

    /**
     * 普通文本编码处理
     *
     * @param paragraphText
     * @param param
     * @return
     */
    private static String textHandler(String paragraphText, Map<String, Object> param) {
        log.info("文本编码处理开始:{}", paragraphText);
        String reg = "\\$\\{[A-Z]+}";
        Pattern compile = Pattern.compile(reg);
        for (Map.Entry<String, Object> entry : param.entrySet()) {
            Matcher matcher = compile.matcher(entry.getKey());
            if (matcher.find()) {
                paragraphText = paragraphText.replace(entry.getKey(), entry.getValue().toString());
            }
        }
        log.info("文本编码处理结束:{}", paragraphText);
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
        if (!"docx".equalsIgnoreCase(docType)) {
            throw new IllegalArgumentException("不是支持的docx类型");
        }
    }
}
