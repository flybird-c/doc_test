package com.kedacom.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.springframework.http.MediaType;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;

/**
 * @version 1.1.0
 */
@Slf4j
public class DocUtilv2 {

    private static final String WINGDINGS_SQUARE_TURE_FLAG = "wingdings_square_ture";

    private static final String WINGDINGS_SQUARE_FALSE_FLAG = "wingdings_square_false";

    private static final String FS_PIC_PREFIX = "/var/ftphome/kvms3000/fs/openImage/";

    private static final String FS_NET_PIC_PREFIX = "/var/ftphome/kvms3000http";

    private static final String FS_PIC_LOCAL_PREFIX = "/var/ftphome/kvms3000/files/uploads/images/";

    private static final String DOCX_TYPE = "docx";

    private static final String PDF_TYPE_POSTFIX = ".pdf";

    private static final String HTML_TYPE_POSTFIX = ".html";

    private static final String FILE_POINT = ".";

    private static final String WINGDINGS_SQUARE = "Wingdings 2";

    private static final String WINGDINGS_SQUARE_TURE = "\u0052";

    private static final String WINGDINGS_SQUARE_FALSE = "\u0030";

    private DocUtilv2() {
    }

    /**
     * 获取文档中包含的编码
     * @param codeList 字典中的编码集
     * @param srcWordPath word文档的全路径
     * @return List 文档中包含的编码
     */
    public static List<String> getDocCodes(List<String> codeList , String srcWordPath){
        List<String> codes = new ArrayList<>();
        if (CollectionUtils.isEmpty(codeList)) {
            log.error("编码集不能为空");
            return codes;
        }
        String docType = getFileExtension(srcWordPath);
        if (!docType.equalsIgnoreCase(DOCX_TYPE)){
            log.error("不支持的文件类型： " + docType);
            return codes;
        }
        try(FileInputStream fis = new FileInputStream(srcWordPath)){
            CustomXWPFDocument document  = new CustomXWPFDocument(fis);
            // 获取页眉中的编码
            getCodesInHeader(codes,codeList,document);
            // 获取页脚中的编码
            getCodesInFooter(codes,codeList,document);
            // 获取段落中的编码
            getCodesInParagraph(codes,codeList,document);
            // 获取表格中的编码
            getCodesInTable(codes,codeList,document);
        }catch(Exception e){
            log.info("【获取编码】异常:" + e);
        }
        return codes;
    }

    /**
     * 替换编码
     * @param param key-需要替换的编码，value-需要替换的值
     * @param srcWordPath word文档的全路径
     * @return 生成的文档全路径
     */
    public static String replaceWordCode(Map<String,Object> param , String srcWordPath){
        if (CollectionUtils.isEmpty(param)) {
            log.error("参数不能为空");
            return null;
        }
        String docType = getFileExtension(srcWordPath);
        if (!docType.equalsIgnoreCase(DOCX_TYPE)){
            log.error("不支持{}文件类型",docType);
            return null;
        }
        String fileName = getFileName(srcWordPath);
        String targetPath = fileName + (new Date()).getTime() + FILE_POINT + docType;
        File targetPathFile = new File(targetPath);
        try(FileOutputStream fos = new FileOutputStream(targetPathFile);
            FileInputStream fis = new FileInputStream(srcWordPath)){
            CustomXWPFDocument document  = new CustomXWPFDocument(fis);
            List<XWPFParagraph> paragraphList = document.getParagraphs();
            if(CollectionUtils.isEmpty(paragraphList)){
                return null;
            }
            // 处理段落中图片编码itemParam
            Map<String, Object> itemParam = handItemCodes(param);
            itemParam.putAll(param);
            // 处理段落中特殊编码singParam
            Map<String, Object> singParam = new HashMap<>();
            handleSingleCode(singParam,param,paragraphList);
            // 处理表格中特殊编码
            handTableCodes(singParam,param,document);
            param.putAll(singParam);
            // 替换模板
            replaceCodeInWord(itemParam,param,document);
            document.write(fos);
        }catch(Exception e){
            log.error("替换模板异常{}",srcWordPath);
            e.printStackTrace();
        }
        return targetPath;
    }

    private static void replaceCodeInWord(Map<String,Object> itemParam,Map<String,Object> param,
                                          CustomXWPFDocument document) throws FileNotFoundException, InvalidFormatException {
        List<XWPFParagraph> paragraphList = document.getParagraphs();
        for(XWPFParagraph paragraph:paragraphList){
            replaceImageInParagraph(paragraph,document,itemParam);
            replaceWordInParagraph(paragraph,param);
        }
        creatTableInParagraph(document,param);
        replaceTableInParagraph(document,param,itemParam);
        replaceCodeInFooter(document,param);
        replaceCodeInHeader(document,param);
    }

    public static Map<String, Object> handItemCodes(Map<String,Object> param){
        Object item = param.get("ITEM");
        Map<String, Object> itemMap = new HashMap<>();
        if (!ObjectUtils.isEmpty(item)) {
            for (Map.Entry<String, Object> entry : param.entrySet()) {
                itemMap.put(entry.getKey() + "-" + item , entry.getValue());
            }
        }
        return itemMap;
    }

    private static void handTableCodes(Map<String,Object> singParam,
                                       Map<String,Object> param ,CustomXWPFDocument document){
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells) {
                    List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                    handleSingleCode(singParam,param,paragraphs);
                }
            }
        }
    }


    private static void replaceCodeInFooter(CustomXWPFDocument document, Map<String, Object> param) {
        // 获取页脚中的编码
        List<XWPFFooter> footerList = document.getFooterList();
        for (XWPFFooter xwpfFooter : footerList) {
            List<XWPFParagraph> listParagraph = xwpfFooter.getListParagraph();
            for (XWPFParagraph xwpfParagraph : listParagraph) {
                //获取文字编码
                replaceWordInParagraph(xwpfParagraph,param);
            }
        }
    }

    private static void replaceCodeInHeader(CustomXWPFDocument document, Map<String, Object> param) {
        // 获取页眉中的编码
        List<XWPFHeader> headerList = document.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            List<XWPFParagraph> listParagraph = xwpfHeader.getListParagraph();
            for (XWPFParagraph xwpfParagraph : listParagraph) {
                //获取文字编码
                replaceWordInParagraph(xwpfParagraph,param);
            }
        }
    }


    private static void handleSingleCode(Map<String, Object> hashMap , Map<String, Object> param, List<XWPFParagraph> paragraphList){
        //替换之前处理“_”的编码 ，先读取模板中的编码，再进行处理
        for (Map.Entry<String, Object> entry : param.entrySet()) {
            String key = entry.getKey();
            for (XWPFParagraph xwpfParagraph : paragraphList) {
                //获取所有runs去掉空格拼接成整段文本(一个段落)
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                StringBuilder stringBuilder = new StringBuilder();
                for (XWPFRun run : runs) {
                    stringBuilder.append(run.toString().trim());
                }
                String runText = stringBuilder.toString();
                String[] split1 = runText.split("\\$");
                for (String s : split1) {
                    if(s.startsWith("{")){
                        int i = s.indexOf("}");
                        String code = s.substring(1, i);
                        log.debug("【获取文字编码】:" + code);
                        //针对勾选框和item特殊编码处理,将所有映射传入hashmap里
                        replaceSpecialWordCode(hashMap,param,key,code);
                    }
                }
            }
        }
        // 处理勾选框编码
        handNormalSquareCodes(hashMap,param);
        handItemSquareCodes(hashMap,param);
    }

    private static void replaceSpecialWordCode(Map<String, Object> hashMap,
                                               Map<String, Object> param,
                                               String key, String code){
        // 获取文档中的编码时，RQYY可以识别出RQYY_1,并返回RQYY
        if (!code.contains("_") && !code.contains("-") ) {
            return;
        }
        replaceSquareWordCode(hashMap,key,code);
        replaceItemWordCode(hashMap,param,key,code);
    }

    private static void handNormalSquareCodes(Map<String,Object> singParam,
                                              Map<String,Object> param){
        for (Map.Entry<String, Object> entry : param.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();
            if (!ObjectUtils.isEmpty(value)){
                //这里是针对勾选框做判断,如果包含逗号则为多选框,这里的第一个判断是多余的
                if (!value.toString().contains(",")){
                    singParam.put(key + "_" + value,WINGDINGS_SQUARE_TURE_FLAG);
                }else {
                    String[] split1 = value.toString().split(",");
                    for (String s1 : split1) {
                        singParam.put(key + "_" + s1,WINGDINGS_SQUARE_TURE_FLAG);
                    }
                }
            }
        }
    }

    private static void handItemSquareCodes(Map<String,Object> singParam,
                                            Map<String,Object> param){
        //处理“XY_血液-2”形式的编码
        for (String key : singParam.keySet()) {
            if (key.contains("_") && key.contains("-")){
                //s1 "XY , 血液-2"
                String[] s1 = key.split("_");
                //s2  "血液-2"
                String s2 = s1[1];
                //s3   "血液，2"
                String[] s3 = s2.split("-");
                if (param.get(s1[0]).equals(s3[0])){
                    if (!s3[1].equals(param.get("ITEM"))){
                        singParam.put(key,WINGDINGS_SQUARE_FALSE_FLAG);
                    }else{
                        singParam.put(key,WINGDINGS_SQUARE_TURE_FLAG);
                    }
                }
            }
        }
    }

    private static void replaceSquareWordCode(Map<String, Object> hashMap,
                                              String key, String code){
        String result = null;
        //判断带_的勾选框编码,,如果符合条件则将把key加入到map里
        if (code.contains("_")) {
            String[] s1 = code.split("_");
            if (!ObjectUtils.isEmpty(s1)) {
                if (s1.length == 3){
                    result = s1[0] + "_" + s1[1];
                }else {
                    result = s1[0];
                }
            }
            if (key.equals(result)){
                hashMap.put(code,WINGDINGS_SQUARE_FALSE_FLAG);
            }
        }
    }

    private static void replaceItemWordCode(Map<String, Object> hashMap,
                                            Map<String, Object> param,
                                            String key, String code){
        String result;
        Object item = null ;
        for (String s : param.keySet()) {
            if (s.startsWith("ITEM")){
                item = param.get(s);
            }
        }
        if (!ObjectUtils.isEmpty(item) && code.contains("-")){
            String[] s2 = code.split("-");
            if (!ObjectUtils.isEmpty(s2)) {
                result = s2[0];
                if (key.equals(result)) {
                    if (item.toString().equals(s2[s2.length - 1])) {
                        hashMap.put(code, param.get(result));
                    } else {
                        hashMap.put(code, "\u3000\u3000");
                    }
                }
            }
        }
    }

    private static String handleSpecialWordCode(String code){
        String result = null;
        // 获取文档中的编码时，SZNF可以识别出SZNF-1,SZNF
        if (code.contains("-")) {
            String[] s1 = code.split("-");
            if (!ObjectUtils.isEmpty(s1)) {
                result = s1[0];
            }
        }

        // 获取文档中的编码时，RQYY可以识别出RQYY_1,并返回RQYY,处理“XY_血液-2”形式的编码
        if (code.contains("_")) {
            String[] s1 = code.split("_");
            if (!ObjectUtils.isEmpty(s1)) {
                result = s1[0];
                if (s1.length == 3){
                    result = s1[0] + "_" + s1[1];
                }
            }
        }
        return result;
    }


    private static void replaceWordInParagraph(XWPFParagraph paragraph, Map<String, Object> param){
        Object item = param.get("ITEM");
        if (!ObjectUtils.isEmpty(item)){
            param.put("ITEM","");
        }
        String text = paragraph.getText();
        if(StringUtils.isEmpty(text)){
            return;
        }
        for (Map.Entry<String, Object> entry : param.entrySet()) {
            String key = entry.getKey();
            Object value = entry.getValue();
            //替换
            if(!text.contains(key)){
                continue;
            }
            List<XWPFRun> runs = paragraph.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                log.debug("run:" + run.toString());
                if (!run.toString().contains("$")){
                    continue;
                }
                // 包含编码的run
                String s = run.toString();
                UnderlinePatterns underline = run.getUnderline();
                boolean bold = run.isBold();
                int fontSize = run.getFontSize();
                String fontFamily1 = run.getFontFamily();
                List<Integer> removeFlag = new ArrayList<>();
                // 处理${XXXXX}编码不在同一个run中的情况，i为当前runs循环层数,这里大概是想拼接不同行的runs类内容,然后整合判断
                String paragraphText = splitRunText(removeFlag,runs,i);
                String dollarCode = splitRealCode(key);
                if (paragraphText.contains(dollarCode)){
                    log.info("paragraphText: " + paragraphText);
                    //  处理编码所在run前含有编码和字符的情况
                    String result = handleExtraRunAroundCode(paragraphText, key, value);
                    String runText = s.replace(s,result);
                    runText = replaceRepeatCodeInRun(runText,key,dollarCode,value);//之前的处理步骤只处理了一个编码变为flag,这里将所有相同的${xxx}编码替换成flag
                    //这里删除跨行的run插入新的一行run
                    removeExtraRun(paragraph,removeFlag);
                    XWPFRun xwpfRun = paragraph.insertNewRun(i);//todo 这里直接在循环中插入新的run会引起异常么,还是会遍历这个新的run
                    xwpfRun.setUnderline(underline);
                    xwpfRun.setBold(bold);
                    if (fontSize != -1) {
                        xwpfRun.setFontSize(fontSize);
                    }
                    xwpfRun.setFontFamily(fontFamily1);
                    replaceWordInRun(paragraph,xwpfRun,key,value,runText);
                }
            }
        }
    }

    private static String replaceRepeatCodeInRun(String runText,String key,String key1,Object value){
        //处理一个run中含有重复的编码，替换不掉的问题
        if (runText.contains(key)){
            runText = runText.replace(key1,value.toString());
        }
        return runText;
    }


    private static void removeExtraRun(XWPFParagraph paragraph,List<Integer> removeFlag){
        //paragraph移除run时会自动补位，所以移除第一run出现的位置多次
        for (int i1 = 0; i1 < removeFlag.size(); i1++) {
            paragraph.removeRun(removeFlag.get(0));
        }
    }

    private static String handleExtraRunAroundCode(String paragraphText, String key , Object value){
        String[] split1 = paragraphText.split("\\$");
        if (!ObjectUtils.isEmpty(split1)){
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.append(split1[0]);//增加前面的字符
            if (split1.length == 1){
                stringBuilder.append(value);//如果长度等于1直接将flag标签加到之前的字符后边
            }
            for (int i1 = 1; i1 < split1.length; i1++) {
                if (!split1[i1].contains(key)){//如果当前循环的文本不包含key,则将用来拆分的$原样增加到sb里
                    stringBuilder.append("$" + split1[i1]);
                }else {
                    stringBuilder.append(value);
                    String[] split = split1[i1].split("}");
                    if (!ObjectUtils.isEmpty(split) && split.length != 1){
                        stringBuilder.append(split[split.length - 1]);
                    }
                    for (int i2 = i1 + 1; i2 < split1.length; i2++) {
                        stringBuilder.append("$" + split1[i2]);
                    }
                    break;
                }
            }
            value = stringBuilder.toString() ;
        }
        value = value.toString().trim();
        return value.toString();
    }

    private static String splitRunText(List<Integer> removeFlag , List<XWPFRun> runs , int i){
        StringBuilder stringBuilder = new StringBuilder();
        //拼接了一个段落的所有内容
        for (int j = i; j < runs.size(); j++) {
            String s1 = runs.get(j).toString();
            stringBuilder.append(s1);
            removeFlag.add(j);
            if (s1.endsWith("}")){
                break;
            }
        }
        return stringBuilder.toString();
    }

    private static String splitRealCode(String key){
        StringBuilder stringBuilder = new StringBuilder();
        stringBuilder.append("${");
        stringBuilder.append(key);
        stringBuilder.append("}");
        return stringBuilder.toString();
    }



    private static void replaceWordInRun(XWPFParagraph paragraph,XWPFRun xwpfRun , String key , Object value , String runText){
        if (runText.equals(WINGDINGS_SQUARE_TURE_FLAG)) {
            xwpfRun.setText(WINGDINGS_SQUARE_TURE);
            xwpfRun.setFontFamily(WINGDINGS_SQUARE);
        }else if (runText.equals(WINGDINGS_SQUARE_FALSE_FLAG)) {
            xwpfRun.setText(WINGDINGS_SQUARE_FALSE);
            xwpfRun.setFontFamily(WINGDINGS_SQUARE);
        }else if ((runText.contains(WINGDINGS_SQUARE_TURE_FLAG)) ) {
            String[] rs = runText.split(WINGDINGS_SQUARE_TURE_FLAG);
            if (rs.length > 1) {
                xwpfRun.setText(WINGDINGS_SQUARE_TURE);
                xwpfRun.setFontFamily(WINGDINGS_SQUARE);
                XWPFRun xwpfRun1 = paragraph.createRun();//这里和外边的插入insertNewRun方法有什么区别,插入后在创建一个新的run是在文件末尾么
                xwpfRun1.setUnderline(xwpfRun.getUnderline());
                xwpfRun1.setBold(xwpfRun.isBold());
                xwpfRun1.setText(rs[1]);
                paragraph.addRun(xwpfRun1);
            }
        }else {
            xwpfRun.setText(runText);
        }
        if (key.startsWith("_") && ObjectUtils.isEmpty(value)){
            xwpfRun.setStrikeThrough(true);
            xwpfRun.setText("\u3000\u3000\u3000\u3000\u3000\u3000");
        }
        paragraph.addRun(xwpfRun);
    }


    private static void replaceImageInParagraph(XWPFParagraph paragraph , CustomXWPFDocument document ,
                                                Map<String, Object> param) throws FileNotFoundException, InvalidFormatException {
        List<XWPFRun> runList = paragraph.getRuns();
        for (int i = 0; i < runList.size(); i++) {
            List<XWPFPicture> pictures = runList.get(i).getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                String desc = picture.getDescription();
                if (desc.startsWith("$")){
                    desc = desc.substring(desc.indexOf("{") + 1, desc.indexOf("}"));
                    if (param.keySet().contains(desc)){
                        log.debug("【图片编码】:" + desc);
                        CTPicture ctPicture = picture.getCTPicture();
                        String picPaths = (String) param.get(desc);
                        //图片编码对应的值为空时不进行处理
                        if(StringUtils.isEmpty(picPaths)){
                            continue;
                        }
                        replaceImageInParagraph(picPaths,ctPicture,paragraph,document,i,desc);
                    }
                }
            }
        }
    }

    private static void replaceImageInParagraph(String picPaths,CTPicture ctPicture,XWPFParagraph paragraph,
                                                CustomXWPFDocument document,int i,String desc)
            throws FileNotFoundException, InvalidFormatException {
        //多张图片处理
        String[] split1 = picPaths.split(";");
        if (split1.length == 1){
            for (String picPath : split1) {
                String blipId = handleBlipId(picPath, document);
                if (StringUtils.isEmpty(blipId)){
                    continue;
                }
                ctPicture.getBlipFill().getBlip().setEmbed(blipId);
            }
        }else {
            replaceMultipleImageInParagraph(split1,ctPicture,paragraph,document,i,desc);
        }

    }

    private static void replaceMultipleImageInParagraph(String[] split1,CTPicture ctPicture,XWPFParagraph paragraph,
                                                        CustomXWPFDocument document,int i,String desc)
            throws FileNotFoundException, InvalidFormatException {
        //图片id
        long id = ctPicture.getNvPicPr().getCNvPr().getId();
        //图片的长度
        long cx = ctPicture.getSpPr().getXfrm().getExt().getCx();
        //图片的宽度
        long cy = ctPicture.getSpPr().getXfrm().getExt().getCy();
        paragraph.removeRun(i);
        for (String picPath : split1) {
            String blipId = handleBlipId(picPath, document);
            if (StringUtils.isEmpty(blipId)){
                continue;
            }
            document.createPicture(blipId, (int) id, cx, cy, desc, paragraph,i);
        }

    }



    private static String handleBlipId(String picPath, CustomXWPFDocument document) throws FileNotFoundException, InvalidFormatException {
        //判断picPath路径  /var/ftphome/kvms3000/fs/openImage/20220728095420850_0.jpg
        //拼接成           /var/ftphome/kvms3000/files/uploads/images/20220728/20220728095420850_0.jpg
        //拼接成           /var/ftphome/kvms3000http://172.16.231.169:19505/fs/openImage/20220817090913763_0.JPG
        String[] split = picPath.split("\\.");
        String picType = split[split.length - 1];
        File file = new File(picPath);
        if(picPath.startsWith(FS_PIC_PREFIX)){
            String[] split2 = picPath.split("/");
            String picName = split2[split2.length - 1];
            String subPicPath = picName.substring(0, 8);
            picPath = FS_PIC_LOCAL_PREFIX + subPicPath + File.separator + picName;
        }
        if(picPath.startsWith(FS_NET_PIC_PREFIX)){
            String picNetUrl = picPath.substring(21);
            if (!verifyImage(getPicInputStream(picNetUrl))) {
                log.error("获取网络地址流信息失败");
                return null;
            }
            InputStream picInputStream = getPicInputStream(picNetUrl);
            if (null != picInputStream){
                return document.addPictureData(getPicInputStream(picNetUrl), getPictureType(picType));
            }
        }
        if (!verifyImage(new FileInputStream(file))) {
            log.info("图片损坏： " + picPath);
            return null;
        }
        return document.addPictureData(new FileInputStream(file), getPictureType(picType));
    }

    private static boolean verifyImage(InputStream inputStream) {
        try (ImageInputStream iis = ImageIO.createImageInputStream(inputStream)) {
            Iterator<ImageReader> iter = ImageIO.getImageReaders(iis);
            if (!iter.hasNext()) {
                log.error("No readers found!");
                return false;
            }
            ImageReader reader = iter.next();
            reader.setInput(iis, true);
            reader.getWidth(0);
            return true;
        } catch (Exception e) {
            log.error(e.getMessage());
            return false;
        }
    }


    private static InputStream getPicInputStream(String url){
        try {
            URL httpUrl = new URL(url);
            HttpURLConnection conn =(HttpURLConnection) httpUrl.openConnection();
            conn.setDoInput(true);
            conn.setDoOutput(true);
            //避免post方式缓存报错
            conn.setUseCaches(false);
            //连接指定资源
            conn.connect();
            //获取文件二进制流
            return conn.getInputStream();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }



    private static void replaceTableInParagraph(CustomXWPFDocument document , Map<String, Object> param
            , Map<String, Object> hashMapPic) throws FileNotFoundException, InvalidFormatException {
        Iterator<XWPFTable> tablesIterator = document.getTablesIterator();
        if (!tablesIterator.hasNext()){
            return;
        }
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            XWPFTable table = tables.get(i);
            List<XWPFTableRow> rows = table.getRows();
            for (XWPFTableRow row : rows) {
                List<XWPFTableCell> tableCells = row.getTableCells();
                for (XWPFTableCell tableCell : tableCells) {
                    List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                    for (XWPFParagraph paragraph : paragraphs) {
                        replaceWordInParagraph(paragraph,param);
                        replaceImageInParagraph(paragraph,document,hashMapPic);
                    }
                }
            }
        }
    }

    private static void creatTableInParagraph(CustomXWPFDocument document , Map<String, Object> param){
        Iterator<XWPFTable> tablesIterator = document.getTablesIterator();
        if (!tablesIterator.hasNext()){
            return;
        }
        List<XWPFTable> tables = document.getTables();
        for (int i = 0; i < tables.size(); i++) {
            //只处理行数大于等于2的表格，且不循环表头
            XWPFTable table = tables.get(i);
            CTTblPr tblPr = table.getCTTbl().getTblPr();
            Set<String> strings = param.keySet();
            for (String string : strings) {
                if (!tblPr.toString().contains(string)){
                    continue;
                }
                Object o = param.get(string);
                if (o instanceof ArrayList<?>){
                    List<List<Object>> lists = handleTableDates(o);
                    insertTableByCode(lists,string,table);
                }
            }
        }

    }


    private static void insertTableByCode(List<List<Object>> lists , String string , XWPFTable table){
        if (!CollectionUtils.isEmpty(lists)){
            if(string.equals("SSCWJCJL")){
                insertTable(table, lists, 5 , 11 , 1,1);
            } else  if(string.equals("YHDJL")){
                insertTable(table, lists, 11 , 2 , 1,1);
            }else {
                insertTable(table, lists, table.getNumberOfRows() - 2 , 2 , 1,0);
            }
        }
    }




    private static void insertTable(XWPFTable table,List<List<Object>> daList , Integer nowTableSize ,
                                    Integer insertTablePos , Integer insertDatePos , Integer flag){
        addNewTableRow(table,nowTableSize,daList,insertDatePos);
        //创建行,根据需要插入的数据添加新行，不处理表头
        for(int i = 0; i < daList.size(); i++){
            List<XWPFTableCell> cells = table.getRow(i + insertTablePos - insertDatePos).getTableCells();
            for(int j = 0; j < cells.size(); j++){
                XWPFTableCell tableCell = cells.get(j);
                String insertDate = daList.get(i).get(j).toString() ;
                List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                insertTableDate(paragraphs,insertDate,flag);
            }
        }
    }

    private static void insertTableDate(List<XWPFParagraph> paragraphs,String insertDate,Integer flag){
        for (XWPFParagraph paragraph : paragraphs) {
            // 如果run为空，判定新插入的数据，创建新的run,设置字体样式
            List<XWPFRun> runs = paragraph.getRuns();
            if (ObjectUtils.isEmpty(runs)){
                handleInsertTableDate(paragraph,insertDate,flag);
                // 复制出来，新插入的行，runs不为空
            }else {
                for (int i1 = 0; i1 < runs.size(); i1++) {
                    paragraph.removeRun(i1);
                    handleInsertTableDate(paragraph,insertDate,flag);
                }
            }
        }
    }

    private static void addNewTableRow(XWPFTable table, Integer nowTableSize, List<List<Object>> daList,Integer insertTablePos){
        log.info("现有文档中的表格行数: " + nowTableSize );
        //实际需要插入的行数
        int size1 = daList.size();
        log.info("实际需要插入的行数: " + size1);

        if (size1 > nowTableSize ){
            for(int i = 0; i < size1 - nowTableSize; i++){
                //添加一个新行
                XWPFTableRow targetRow= table.insertNewTableRow(insertTablePos);
                XWPFTableRow sourceRow = table.getRow(insertTablePos - 1);
                targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
                List<XWPFTableCell> tableCells = sourceRow.getTableCells();
                if (CollectionUtils.isEmpty(tableCells)) {
                    return;
                }
                createCellsAndCopyStyles(targetRow,sourceRow);
            }
        }
    }

    private static void handleInsertTableDate(XWPFParagraph paragraph , String insertDate, Integer flag){
        XWPFRun run = paragraph.createRun();
        // 处理清单中的字体
        if (flag.equals(0)) {
            // 4号字体 大小 14
            run.setFontSize(14);
            run.setText(insertDate);
            run.setFontFamily("仿宋_GB2312");
            //处理登记表中的字体
        }else if (flag.equals(1)){
            // 小4字体 大小 12
            run.setFontSize(12);
            run.setText(insertDate);
            run.setFontFamily("仿宋");
        }
    }

    private static void createCellsAndCopyStyles(XWPFTableRow targetRow, XWPFTableRow sourceRow) {
        List<XWPFTableCell> tableCells = sourceRow.getTableCells();
        for (XWPFTableCell sourceCell : tableCells) {
            XWPFTableCell newCell = targetRow.addNewTableCell();
            newCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
            List<XWPFParagraph> sourceParagraphs = sourceCell.getParagraphs();
            if (CollectionUtils.isEmpty(sourceParagraphs)) {
                continue;
            }
            XWPFParagraph sourceParagraph = sourceParagraphs.get(0);
            List<XWPFParagraph> targetParagraphs = newCell.getParagraphs();
            if (CollectionUtils.isEmpty(targetParagraphs)) {
                XWPFParagraph p = newCell.addParagraph();
                p.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                XWPFRun run = p.getRuns().isEmpty() ? p.createRun() : p.getRuns().get(0);
                run.setFontFamily(sourceParagraph.getRuns().get(0).getFontFamily());
            } else {
                XWPFParagraph p = targetParagraphs.get(0);
                p.getCTP().setPPr(sourceParagraph.getCTP().getPPr());
                XWPFRun run = p.getRuns().isEmpty() ? p.createRun() : p.getRuns().get(0);
                List<XWPFRun> runs = sourceParagraph.getRuns();
                if (!CollectionUtils.isEmpty(runs)){
                    run.setFontFamily(runs.get(0).getFontFamily());
                }
            }
        }
    }


    private static void getCodesInHeader(List<String> codes, List<String> codeList, CustomXWPFDocument document){
        // 获取页眉中的编码
        List<XWPFHeader> headerList = document.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            List<XWPFParagraph> listParagraph = xwpfHeader.getListParagraph();
            for (XWPFParagraph xwpfParagraph : listParagraph) {
                //获取文字编码
                getWordCodes(codes,codeList,xwpfParagraph);
            }
        }
    }

    private static void getCodesInFooter(List<String> codes, List<String> codeList, CustomXWPFDocument document){
        // 获取页眉中的编码
        List<XWPFFooter> footerList = document.getFooterList();
        for (XWPFFooter xwpfFooter : footerList) {
            List<XWPFParagraph> listParagraph = xwpfFooter.getListParagraph();
            for (XWPFParagraph xwpfParagraph : listParagraph) {
                //获取文字编码
                getWordCodes(codes,codeList,xwpfParagraph);
            }
        }
    }

    private static void getCodesInParagraph(List<String> codes, List<String> codeList, CustomXWPFDocument document){
        List<XWPFParagraph> paragraphList = document.getParagraphs();
        if(!CollectionUtils.isEmpty(paragraphList)){
            // 获取文档中的编码
            for(XWPFParagraph paragraph:paragraphList){
                //获取图片编码
                getPicCodes(codes,codeList,paragraph);
                //获取文字编码
                getWordCodes(codes,codeList,paragraph);
            }
        }
    }

    private static void getCodesInTable(List<String> codes, List<String> codeList, CustomXWPFDocument document){
        Iterator<XWPFTable> tablesIterator = document.getTablesIterator();
        if (tablesIterator.hasNext()){
            List<XWPFTable> tables = document.getTables();
            tables.forEach(table->{
                //获取表格的编码
                getTableCodInParagraph(codes, codeList,table);
                //获取表格中的文字编码
                getWordCodesInTable(codes, codeList,table);
            });
        }
    }

    private static void getTableCodInParagraph(List<String> codes, List<String> codeList,XWPFTable table){
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (!ObjectUtils.isEmpty(tblPr)) {
            for (String code : codeList) {
                if (tblPr.toString().equals(code)) {
                    log.info("【获取表格编码】:" + tblPr.toString());
                    codes.add(code);
                }
            }
        }
    }

    private static void getWordCodesInTable(List<String> codes, List<String> codeList,XWPFTable table){
        //获取表格中的文字编码
        List<XWPFTableRow> rows = table.getRows();
        for (XWPFTableRow row : rows) {
            List<XWPFTableCell> tableCells = row.getTableCells();
            for (XWPFTableCell tableCell : tableCells) {
                List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                for (XWPFParagraph paragraph : paragraphs) {
                    getPicCodes(codes, codeList, paragraph);
                    getWordCodes(codes, codeList, paragraph);
                }
            }
        }
    }



    private static void getWordCodes(List<String> codes ,List<String> codeList , XWPFParagraph paragraph){
        List<XWPFRun> runs = paragraph.getRuns();
        StringBuilder stringBuilder = new StringBuilder();
        for (XWPFRun run : runs) {
            stringBuilder.append(run.toString().trim());
        }
        String runText = stringBuilder.toString();
        String[] split = runText.split("\\$");
        for (String s : split) {
            if(s.startsWith("{")){
                int i = s.indexOf("}");
                String code = s.substring(1, i);
                String result = handleSpecialWordCode(code);
                if (codeList.contains(result)){
                    codes.add(result);
                } else if (codeList.contains(code)){
                    codes.add(code);
                }
            }
        }
    }



    public static void getPicCodes(List<String> codes ,List<String> codeList , XWPFParagraph paragraph){
        List<XWPFRun> runList = paragraph.getRuns();
        for (XWPFRun xwpfRun : runList) {
            List<XWPFPicture> pictures = xwpfRun.getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                String picCode = picture.getDescription();
                log.info("【获取图片编码】:" + picCode);
                if (picCode.startsWith("$")){
                    picCode = picCode.substring(picCode.indexOf("{") + 1, picCode.indexOf("}"));
                    String result = handleSpecialWordCode(picCode);
                    if (codeList.contains(result)){
                        codes.add(result);
                    }else if (codeList.contains(picCode)){
                        codes.add(picCode);
                    }
                }
            }
        }
    }

    /**
     * 处理接受的list实体
     * @param obj
     * @return
     */
    private static List<List<Object>> handleTableDates(Object obj) {
        log.info("【处理接受的list实体】 开始" + obj);
        List<List<Object>> allList = new ArrayList<>();
        if (obj == null) {
            return allList;
        }
        if (obj instanceof ArrayList<?>) {
            List<Object> obj1 = (List)obj;
            for (Object o : obj1) {
                if (o instanceof ArrayList<?>){
                    return (List)obj;
                }
                handleTableDates(allList,o);
            }
        }
        log.info("【处理接受的list实体】 结束" + allList);
        return allList;
    }

    private static void handleTableDates(List<List<Object>> allList , Object o) {
        Class<?> aClass = o.getClass();
        List<Object> rowList = new ArrayList<>();
        if (o instanceof JSONArray){
            List<Object> list = JSON.parseArray(o.toString(), Object.class);
            for (Object o1 : list) {
                rowList.add(o1.toString());
            }
            allList.add(rowList);
            return ;
        }
        Field[] fields = aClass.getDeclaredFields();
        try {
            for (Field field : fields) {
                field.setAccessible(true);
                rowList.add(field.get(o).toString());
            }
            allList.add(rowList);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static int getPictureType(String picType){
        int res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PICT;
        if(picType != null){
            if(picType.equalsIgnoreCase("png")){
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }


    /**
     * word转pdf
     * @param filePath word文件全路径
     * @return pdf文件全路径
     */
    public static String getPdfByWordPath(String filePath){
        String filePdfPath = getFileName(filePath);
        String srcPdfPath = filePdfPath + PDF_TYPE_POSTFIX;
        File pdfFile = new File(srcPdfPath);
        if (!pdfFile.exists()){
            try(FileOutputStream os = new FileOutputStream(pdfFile)) {
                Document doc = new Document(filePath);
                doc.save(os, SaveFormat.PDF);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return srcPdfPath;
    }

    /**
     * word转html
     * @param filePath word文件全路径
     * @return html文件全路径
     */
    public static String getHtmlByWordPath(String filePath){
        String filePdfPath = getFileName(filePath);
        String targetHtmlPath = filePdfPath + HTML_TYPE_POSTFIX;
        try(FileInputStream fis = new FileInputStream(filePath);
            FileOutputStream fos = new FileOutputStream(targetHtmlPath);
            OutputStreamWriter osw = new OutputStreamWriter(fos, StandardCharsets.UTF_8)) {
            XWPFDocument doc = new XWPFDocument(fis);
            XHTMLOptions xhtmlOptions = XHTMLOptions.create();
            XHTMLConverter instance = (XHTMLConverter)XHTMLConverter.getInstance();
            instance.convert(doc,osw,xhtmlOptions);
            return targetHtmlPath;
        } catch (IOException e){
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 打开pdf文档
     * @param fileLocalName pdf文档全路径
     * @param response
     */
    public static void openPdf(String fileLocalName, HttpServletResponse response) {
        getFileResponse(fileLocalName,MediaType.APPLICATION_PDF_VALUE,response);
    }


    /**
     * 打开html文档
     * @param fileLocalName html文档全路径
     * @param response
     */
    public static void openHtml(String fileLocalName, HttpServletResponse response) {
        getFileResponse(fileLocalName, MediaType.TEXT_HTML_VALUE,response);
    }

    public static void getFileResponse(String fileName, String fileType, HttpServletResponse response){
        File file = new File(fileName);
        response.setContentType(fileType);
        try {
            response.addHeader("Content-Disposition", "inline;fileName=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Access-Control-Allow-Origin", "*");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        byte [] buff = new byte[1024*10*10];
        try(FileInputStream input = new FileInputStream(file);
            OutputStream out = response.getOutputStream()) {
            int len ;
            while((len = input.read(buff))>-1){
                out.write(buff,0, len);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 获取文件名
    private static String getFileName (String name) {
        return name.substring(0, name.lastIndexOf(FILE_POINT));
    }

    // 只获取后缀名
    private static String getFileExtension (String name) {
        return name.substring(name.lastIndexOf(FILE_POINT) + 1);
    }

}