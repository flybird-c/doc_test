package com.kedacom.util;

import com.alibaba.fastjson.JSONObject;
import com.aspose.words.Document;
import com.aspose.words.SaveFormat;
import com.kedacom.exception.ServiceException;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.springframework.util.CollectionUtils;
import org.springframework.util.ObjectUtils;
import org.springframework.util.StringUtils;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.net.URLEncoder;
import java.util.*;

@Slf4j
public class DocUtilv1 {

    public static final String WINGDINGS_SQUARE_TURE = "wingdings_square_ture";

    public static final String WINGDINGS_SQUARE_FALSE = "wingdings_square_false";

    /**
     * 获取文档中包含的编码
     * @param codeList 字典中的编码集
     * @param srcWordPath word文档的全路径
     * @return List 文档中包含的编码
     * @throws ServiceException
     */
    public static List getDocCodes(List<String> codeList , String srcWordPath) throws ServiceException {
        List<String> codes = new ArrayList<>();
        String[] split = srcWordPath.split("\\.");
        String docType = split[split.length - 1];
        try{
            if(docType.equalsIgnoreCase("docx")){
                CustomXWPFDocument document  = new CustomXWPFDocument(new FileInputStream(srcWordPath));
                if (CollectionUtils.isEmpty(codeList)) {
                    return codes;
                }
                // 获取页眉中的编码
                List<XWPFHeader> headerList = document.getHeaderList();
                for (XWPFHeader xwpfHeader : headerList) {
                    List<XWPFParagraph> listParagraph = xwpfHeader.getListParagraph();
                    for (XWPFParagraph xwpfParagraph : listParagraph) {
                        //获取文字编码
                        getWordCodes(codes,codeList,xwpfParagraph);
                    }
                }
                // 获取页脚中的编码
                List<XWPFFooter> footerList = document.getFooterList();
                for (XWPFFooter xwpfFooter : footerList) {
                    List<XWPFParagraph> listParagraph = xwpfFooter.getListParagraph();
                    for (XWPFParagraph xwpfParagraph : listParagraph) {
                        //获取文字编码
                        getWordCodes(codes,codeList,xwpfParagraph);
                    }
                }
                List<XWPFParagraph> paragraphList = document.getParagraphs();
                if(CollectionUtils.isEmpty(paragraphList)){
                    return codes;
                }
                for(XWPFParagraph paragraph:paragraphList){
                    //获取图片编码
                    getPicCodes(codes,codeList,paragraph);
                    //获取文字编码
                    getWordCodes(codes,codeList,paragraph);
                }
                //获取表格中的编码
                Iterator<XWPFTable> tablesIterator = document.getTablesIterator();
                if (!tablesIterator.hasNext()){
                    return codes;
                }
                List<XWPFTable> tables = document.getTables();
                for (int i = 0; i < tables.size(); i++) {
                    XWPFTable table = tables.get(i);
                    //获取表格的编码
                    CTTblPr tblPr = table.getCTTbl().getTblPr();
                    if (!ObjectUtils.isEmpty(tblPr) ){
                        for (String code : codeList) {
                            if (tblPr.toString().contains(code)){
                                log.info("【获取表格编码】:" + tblPr.toString());
                                codes.add(code);
                            }
                        }
                    }
                    //获取表格中的文字编码
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        List<XWPFTableCell> tableCells = row.getTableCells();
                        for (XWPFTableCell tableCell : tableCells) {
                            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                            for (XWPFParagraph paragraph : paragraphs) {
                                getPicCodes(codes,codeList,paragraph);
                                getWordCodes(codes,codeList,paragraph);
                            }
                        }
                    }
                }
            }

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
     * @throws ServiceException
     */
    public static String replaceWordCode(Map<String,Object> param , String srcWordPath) throws ServiceException, IOException {
        String[] split = srcWordPath.split("\\.");
        String filePaths = split[0];
        String docType = split[split.length - 1];
        String targetPath = filePaths + (new Date()).getTime() + "." + docType;
        FileOutputStream fos = new FileOutputStream(new File(targetPath));
        try{
            if(docType.equalsIgnoreCase("docx")){
                CustomXWPFDocument document  = new CustomXWPFDocument(new FileInputStream(srcWordPath));
                if (CollectionUtils.isEmpty(param)) {
                    return null;
                }
                List<XWPFParagraph> paragraphList = document.getParagraphs();
                if(CollectionUtils.isEmpty(paragraphList)){
                    return null;
                }
                //处理段落选项处理map
                Object item = param.get("ITEM");
                Map<String, Object> hashMapPic = new HashMap<>();
                if (!ObjectUtils.isEmpty(item)) {
                    Set<String> codes = param.keySet();
                    for (String code : codes) {
                        hashMapPic.put(code + "-" + item , param.get(code));
                    }
                }
                hashMapPic.putAll(param);
                //处理勾选框处理map
                Map<String, Object> hashMap = new HashMap<>();
                handleSingleCode(hashMap,param,paragraphList);
                List<XWPFTable> tables = document.getTables();
                for (XWPFTable table : tables) {
                    List<XWPFTableRow> rows = table.getRows();
                    for (XWPFTableRow row : rows) {
                        List<XWPFTableCell> tableCells = row.getTableCells();
                        for (XWPFTableCell tableCell : tableCells) {
                            List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
                            handleSingleCode(hashMap,param,paragraphs);
                        }
                    }
                }
                for (String s : param.keySet()) {
                    Object o = param.get(s);
                    if (!ObjectUtils.isEmpty(o)){
                        if (!o.toString().contains(",")){
                            hashMap.put(s + "_" + o,WINGDINGS_SQUARE_TURE);
                        }else {
                            String[] split1 = o.toString().split("\\,");
                            for (String s1 : split1) {
                                hashMap.put(s + "_" + s1,WINGDINGS_SQUARE_TURE);
                            }
                        }
                    }
                }
                //处理“XY_血液-2”形式的编码
                for (String s : hashMap.keySet()) {
                    if (!s.contains("_")){
                        continue;
                    }
                    if (!s.contains("-")){
                        continue;
                    }
                    //s1 "XY , 血液-2"
                    String[] s1 = s.split("_");
                    //s2  "血液-2"
                    String s2 = s1[1];
                    //s3   "血液，2"
                    String[] s3 = s2.split("-");
                    if (!param.get(s1[0]).equals(s3[0])){
                        continue;
                    }
                    if (!s3[1].equals(param.get("ITEM"))){
                        hashMap.put(s,WINGDINGS_SQUARE_FALSE);
                    }else{
                        hashMap.put(s,WINGDINGS_SQUARE_TURE);
                    }
                }
                param.putAll(hashMap);
                for(XWPFParagraph paragraph:paragraphList){
                    replaceImageInParagraph(paragraph,document,hashMapPic);
                    replaceWordInParagraph(paragraph,param);
                }
                creatTableInParagraph(document,param);
                replaceTableInParagraph(document,param,hashMapPic);
//                setWholeFontsStyle(document);
                replaceCodeInFooter(document,param);
                replaceCodeInHeader(document,param);
                document.write(fos);
            }
            fos.flush();
            fos.close();
        }catch(Exception e){
            e.printStackTrace();
            fos.flush();
            fos.close();
        }
        return targetPath;
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

    // 设置全局字体样式，但无法指定字体大小
    private static void setWholeFontsStyle(CustomXWPFDocument document){
        XWPFStyles styles = document.getStyles();
        CTFonts fonts = CTFonts.Factory.newInstance();
        fonts.setAscii("仿宋_GB2312");
        fonts.setEastAsia("仿宋_GB2312");
        fonts.setHAnsi("仿宋_GB2312");
        styles.setDefaultFonts(fonts);
        List<XWPFTable> tables1 = document.getTables();
        for (XWPFTable xwpfTable : tables1) {
            String tblStyle = document.getTblStyle(xwpfTable);
            log.info(tblStyle);
        }
    }


    private static void handleSingleCode(Map<String, Object> hashMap , Map<String, Object> param, List<XWPFParagraph> paragraphList){
        //替换之前处理”_“的编码
        Set<String> strings = param.keySet();
        for (String string : strings) {
            for (XWPFParagraph xwpfParagraph : paragraphList) {
                List<XWPFRun> runs = xwpfParagraph.getRuns();
                StringBuffer stringBuffer = new StringBuffer();
                for (XWPFRun run : runs) {
                    stringBuffer.append(run.toString().trim());
                }
                String runText = stringBuffer.toString();
                String[] split1 = runText.split("\\$");
                for (String s : split1) {
                    if(s.startsWith("{")){
                        int i = s.indexOf("}");
                        String code = s.substring(1, i);
                        log.debug("【获取文字编码】:" + code);
                        String code_ = null;
                        // 获取文档中的编码时，RQYY可以识别出RQYY_1,并返回RQYY
                        if (!code.contains("_") && !code.contains("-") ) {
                            continue;
                        }
                        if (code.contains("_")) {
                            String[] s1 = code.split("_");
                            if (!ObjectUtils.isEmpty(s1)) {
                                if (s1.length == 3){
                                    code_ = s1[0] + "_" + s1[1];
                                }else {
                                    code_ = s1[0];
                                }
                            }
                            if (string.equals(code_)){
                                hashMap.put(code,WINGDINGS_SQUARE_FALSE);
                            }
                        }
                        String code_2 = null;
                        Object item = null ;
                        for (String key : param.keySet()) {
                            if (key.startsWith("ITEM")){
                                item = param.get(key);
                            }
                        }
                        if (ObjectUtils.isEmpty(item)){
                            continue;
                        }
                        if (code.contains("-")) {
                            String[] s2 = code.split("-");
                            if (ObjectUtils.isEmpty(s2)) {
                                continue;
                            }
                            code_2 = s2[0];
                            if (string.equals(code_2)) {
                                if (item.toString().equals(s2[s2.length - 1])) {
                                    hashMap.put(code, param.get(code_2));
                                } else {
                                    hashMap.put(code, "\u3000\u3000");
                                }
                            }
                        }
                    }
                }
            }
        }
    }


    /**
     * word转pdf
     * @param filePath word文件全路径
     * @return pdf文件全路径
     * @throws ServiceException
     */
    public static String getPdfByWordPath(String filePath) throws ServiceException {
        String[] split = filePath.split("\\.");
        String filePdfPath = split[0];
        String srcPdfPath = filePdfPath + ".pdf";
        File pdfFile = new File(srcPdfPath);
        if (pdfFile.exists()){
            return srcPdfPath;
        }
        FileOutputStream os = null;
        try {
            os = new FileOutputStream(pdfFile);
            Document doc = new Document(filePath);
            doc.save(os, SaveFormat.PDF);
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if (os != null) {
                try {
                    os.flush();
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return srcPdfPath;
    }

    /**
     * 打开pdf文档
     * @param fileLocalName pdf文档全路径
     * @param response
     * @throws ServiceException
     */
    public static void openPdf(String fileLocalName, HttpServletResponse response) throws ServiceException {
        File file = new File(fileLocalName);
        String fileName = fileLocalName;
        response.setContentType("application/pdf");
        try {
            response.addHeader("Content-Disposition", "inline;fileName=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Access-Control-Allow-Origin", "*");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        byte [] buff = new byte[1024*10*10];
        FileInputStream input = null;
        OutputStream out = null;
        try {
            input = new FileInputStream(file);
            out = response.getOutputStream();
            int len=0;
            while((len=input.read(buff))>-1){
                out.write(buff,0,len);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(input!=null){
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    /**
     * word转html
     * @param filePath word文件全路径
     * @return html文件全路径
     * @throws ServiceException
     */
    public static String getHtmlByWordPath(String filePath) throws ServiceException {
        String[] split = filePath.split("\\.");
        String filePdfPath = split[0];
        String targetHtmlPath = filePdfPath + ".html";
        OutputStreamWriter outputStreamWriter;
        try {
            XWPFDocument doc = new XWPFDocument(new FileInputStream(filePath));
            XHTMLOptions xhtmlOptions = XHTMLOptions.create();
            outputStreamWriter = new OutputStreamWriter(new FileOutputStream(targetHtmlPath),"utf-8");
            XHTMLConverter instance = (XHTMLConverter)XHTMLConverter.getInstance();
            instance.convert(doc,outputStreamWriter,xhtmlOptions);
            readFileContent(targetHtmlPath);
            return targetHtmlPath;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * 打开html文档
     * @param fileLocalName html文档全路径
     * @param response
     * @throws ServiceException
     */
    public static void openHtml(String fileLocalName, HttpServletResponse response) throws ServiceException {
        File file = new File(fileLocalName);
        String fileName = fileLocalName;
        response.setContentType("text/html");
        try {
            response.addHeader("Content-Disposition", "inline;fileName=" + URLEncoder.encode(fileName, "UTF-8"));
            response.setHeader("Access-Control-Allow-Origin", "*");
        } catch (UnsupportedEncodingException e) {
            e.printStackTrace();
        }
        byte [] buff = new byte[1024*10*10];
        FileInputStream input = null;
        OutputStream out = null;
        try {
            input = new FileInputStream(file);
            out = response.getOutputStream();
            int len=0;
            while((len=input.read(buff))>-1){
                out.write(buff,0,len);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(input!=null){
                try {
                    input.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(out!=null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    private static String readFileContent(String fileName){
        try (
                FileInputStream fis = new FileInputStream(fileName);
                BufferedReader reader = new BufferedReader(new InputStreamReader(fis));
        ) {
            String line;
            StringBuffer buffer = new StringBuffer();
            line = reader.readLine();
            while (line != null) {
                buffer.append(line);
                buffer.append("\n");
                line = reader.readLine();
            }
            return buffer.toString();
        } catch (Exception e) {
            e.printStackTrace();
            throw new IllegalStateException(e.getMessage());
        }
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
//            Object value = entry.getValue();
            //替换
            if(!text.contains(key)){
                continue;
            }
            List<XWPFRun> runs = paragraph.getRuns();
            for (int i = 0; i < runs.size(); i++) {
                Object value = entry.getValue();
                XWPFRun run = runs.get(i);
                log.debug("run:" + run.toString());
                if (!run.toString().contains("$")){
                    continue;
                }
                String s = run.toString();
                List<Integer> removeFlag = new ArrayList<>();
                StringBuffer stringBuffer1 = new StringBuffer();
                // 处理${XXXXX}编码不在同一个run中的情况
                for (int j = i; j < runs.size(); j++) {
                    String s1 = runs.get(j).toString();
                    stringBuffer1.append(s1);
                    removeFlag.add(j);
                    if (s1.endsWith("}")){
                        break;
                    }
                }
                UnderlinePatterns underline = run.getUnderline();
                boolean bold = run.isBold();
                int fontSize = run.getFontSize();
                String fontFamily1 = run.getFontFamily();
                StringBuffer stringBuffer = new StringBuffer();
                stringBuffer.append("${");
                stringBuffer.append(key);
                stringBuffer.append("}");
                String key1 = stringBuffer.toString();
                String s1 = stringBuffer1.toString();
                if (!s1.contains(key1)){
                    continue;
                }
                log.info("s1: " + s1);
                //  处理编码所在run前含有编码和字符的情况
                String[] split1 = s1.split("\\$");
                if (!ObjectUtils.isEmpty(split1)){
                    StringBuffer stringBuffer2 = new StringBuffer();
                    stringBuffer2.append(split1[0]);
                    if (split1.length == 1){
                        stringBuffer2.append(value);
                    }
                    for (int i1 = 1; i1 < split1.length; i1++) {
                        if (!split1[i1].contains(key)){
                            stringBuffer2.append("$" + split1[i1]);
                        }else {
                            stringBuffer2.append(value);
                            String[] split = split1[i1].split("}");
                            if (!ObjectUtils.isEmpty(split) && split.length != 1){
                                stringBuffer2.append(split[split.length - 1]);
                            }
                            for (int i2 = i1 + 1; i2 < split1.length; i2++) {
                                stringBuffer2.append("$" + split1[i2]);
                            }
                            break;
                        }
                    }
                    value = stringBuffer2.toString() ;
                }
                value = value.toString().trim();
                String runText = s.replace(s,value.toString());
                //处理一个run中含有重复的编码，替换不掉的问题
                if (runText.contains(key)){
                    runText = runText.replace(key1,param.get(key).toString());
                }
                //paragraph移除run时会自动补位，所以移除第一run出现的位置多次
                for (int i1 = 0; i1 < removeFlag.size(); i1++) {
                    paragraph.removeRun(removeFlag.get(0));
                }
                XWPFRun xwpfRun = paragraph.insertNewRun(i);
                xwpfRun.setUnderline(underline);
                xwpfRun.setBold(bold);
                if (fontSize != -1) {
                    xwpfRun.setFontSize(fontSize);
                }
                xwpfRun.setFontFamily(fontFamily1);
                if (key.contains("GXK")){
                    if ("1".equals(value)) {
                        xwpfRun.setFontFamily("Wingdings 2");
                        xwpfRun.setText("\u0052");
                    }else {
                        xwpfRun.setFontFamily("Wingdings 2");
                        xwpfRun.setText("\u0030");
                    }
                }else {
                    if (runText.equals(WINGDINGS_SQUARE_TURE)) {
                        xwpfRun.setText("\u0052");
                        xwpfRun.setFontFamily("Wingdings 2");
                    }else if (runText.equals(WINGDINGS_SQUARE_FALSE)) {
                        xwpfRun.setText("\u0030");
                        xwpfRun.setFontFamily("Wingdings 2");
                    }else if ((runText.contains(WINGDINGS_SQUARE_TURE)) || (runText.contains(WINGDINGS_SQUARE_FALSE))) {
                        String[] rs ;
                        if (runText.contains(WINGDINGS_SQUARE_TURE)){
                            rs = runText.split(WINGDINGS_SQUARE_TURE);
                        }else {
                            rs = runText.split(WINGDINGS_SQUARE_FALSE);
                        }
                        if (rs.length > 1) {
                            if (runText.contains(WINGDINGS_SQUARE_TURE)){
                                xwpfRun.setText("\u0052");
                            }else {
                                xwpfRun.setText("\u0030");
                            }
                            xwpfRun.setFontFamily("Wingdings 2");
                            XWPFRun xwpfRun1 = paragraph.createRun();
                            xwpfRun1.setUnderline(underline);
                            xwpfRun1.setFontFamily(fontFamily1);
                            xwpfRun1.setBold(bold);
                            xwpfRun1.setText(rs[1]);
                            if (fontSize != -1) {
                                xwpfRun1.setFontSize(fontSize);
                            }
                            paragraph.addRun(xwpfRun1);
                        }
                    }else {
                        xwpfRun.setText(runText);
                    }
                }
                if (key.startsWith("_") && StringUtils.isEmpty(param.get(key))){
                    xwpfRun.setStrike(true);
                    xwpfRun.setText("\u3000\u3000\u3000\u3000\u3000\u3000");
                }
                paragraph.addRun(xwpfRun);
            }
        }
    }


    private static void replaceImageInParagraph(XWPFParagraph paragraph , CustomXWPFDocument document ,
                                                Map<String, Object> param) throws FileNotFoundException, InvalidFormatException {
        List<XWPFRun> runList = paragraph.getRuns();
        for (int i = 0; i < runList.size(); i++) {
            List<XWPFPicture> pictures = runList.get(i).getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                String desc = picture.getDescription();
                if (!desc.startsWith("$")){
                    continue;
                }
                desc = desc.substring(desc.indexOf("{") + 1, desc.indexOf("}"));
                if (!param.keySet().contains(desc)){
                    continue;
                }
                log.debug("【图片编码】:" + desc);
                CTPicture ctPicture = picture.getCTPicture();
                String picPaths = (String) param.get(desc);
                //多张图片处理
                String[] split1 = picPaths.split(";");
                if (split1.length == 1){
                    for (String picPath : split1) {
                        String blipId = getPicBlipIdByPicPath(document,picPath);
                        if (StringUtils.isEmpty(blipId)){
                            continue;
                        }
                        ctPicture.getBlipFill().getBlip().setEmbed(blipId);
                    }
                }else {
                    //图片id
                    long id = ctPicture.getNvPicPr().getCNvPr().getId();
                    //图片的长度
                    long cx = ctPicture.getSpPr().getXfrm().getExt().getCx();
                    //图片的宽度
                    long cy = ctPicture.getSpPr().getXfrm().getExt().getCy();
                    paragraph.removeRun(i);
                    for (String picPath : split1) {
                        String blipId = getPicBlipIdByPicPath(document,picPath);
                        if (StringUtils.isEmpty(blipId)){
                            continue;
                        }
                        document.createPicture(blipId, (int) id, cx, cy, desc, paragraph,i);
                    }
                }
            }

        }
    }

    private static String getPicBlipIdByPicPath(CustomXWPFDocument document ,String picPath) throws FileNotFoundException, InvalidFormatException {
        String[] split = picPath.split("\\.");
        String picType = split[split.length - 1];
        File file = new File(picPath);
        if (!file.exists()){
            log.info("图片不存在： " + picPath);
            return null;
        }
        if (!verifyImage(new FileInputStream(file))) {
            log.info("图片损坏： " + picPath);
            return null;
        }
        String blipId = document.addPictureData(new FileInputStream(file), getPictureType(picType));
        return blipId;
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
            int width = reader.getWidth(0);
            return true;
        } catch (Exception e) {
            log.error(e.getMessage());
            return false;
        }
    }


    private static void replaceTableInParagraph(CustomXWPFDocument document , Map<String, Object> param ,
                                                Map<String, Object> hashMapPic) throws FileNotFoundException, InvalidFormatException {
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
                    List<List> lists = handleTables(o);
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
            }
        }

    }


    private static void insertTable(XWPFTable table,List<List> daList , Integer nowTableSize ,
                                    Integer insertTablePos , Integer insertDatePos , Integer flag){
        log.info("现有文档中的表格行数: " + nowTableSize );
        //实际需要插入的行数
        int size1 = daList.size();
        log.info("实际需要插入的行数: " + size1);

        if (size1 > nowTableSize ){
            for(int i = 0; i < size1 - nowTableSize; i++){
                //添加一个新行
                XWPFTableRow targetRow= table.insertNewTableRow(insertTablePos);
                XWPFTableRow sourceRow = table.getRow(insertTablePos - 1);
                createCellsAndCopyStyles(targetRow,sourceRow );
            }
        }
        //创建行,根据需要插入的数据添加新行，不处理表头
        for(int i = 0; i < daList.size(); i++){
            List<XWPFTableCell> cells = table.getRow(i + insertTablePos - insertDatePos).getTableCells();
            for(int j = 0; j < cells.size(); j++){
                XWPFTableCell tableCell = cells.get(j);
                String insertDate = null ;
                try {
                    insertDate = daList.get(i).get(j).toString();
                } catch (Exception e) {
                    e.printStackTrace();
                }
                List<XWPFParagraph> paragraphs = tableCell.getParagraphs();
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
        targetRow.getCtRow().setTrPr(sourceRow.getCtRow().getTrPr());
        List<XWPFTableCell> tableCells = sourceRow.getTableCells();
        if (CollectionUtils.isEmpty(tableCells)) {
            return;
        }
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


    public static void getWordCodes(List<String> codes ,List<String> codeList , XWPFParagraph paragraph){
        List<XWPFRun> runs = paragraph.getRuns();
        StringBuffer stringBuffer = new StringBuffer();
        for (XWPFRun run : runs) {
            stringBuffer.append(run.toString().trim());
        }
        String runText = stringBuffer.toString();
        String[] split = runText.split("\\$");
        for (String s : split) {
            if(s.startsWith("{")){
                int i = s.indexOf("}");
                String code = s.substring(1, i);
                String code_ = null;

                // 获取文档中的编码时，SZNF可以识别出SZNF-1,SZNF
                if (code.contains("-")) {
                    String[] s1 = code.split("-");
                    if (!ObjectUtils.isEmpty(s1)) {
                        code_ = s1[0];
                    }
                }

                // 获取文档中的编码时，RQYY可以识别出RQYY_1,并返回RQYY,处理“XY_血液-2”形式的编码
                if (code.contains("_")) {
                    String[] s1 = code.split("_");
                    if (!ObjectUtils.isEmpty(s1)) {
                        code_ = s1[0];
                        if (s1.length == 3){
                            code_ = s1[0] + "_" + s1[1];
                        }
                    }
                }

                if (codeList.contains(code_)){
                    codes.add(code_);
                    continue;
                }
                if (codeList.contains(code)){
                    codes.add(code);
                }
            }
        }
    }


    public static void getPicCodes(List<String> codes ,List<String> codeList , XWPFParagraph paragraph){
        List<XWPFRun> runList = paragraph.getRuns();
        for (int i = 0; i < runList.size(); i++) {
            List<XWPFPicture> pictures = runList.get(i).getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                String picCode = picture.getDescription();
                log.info("【获取图片编码】:" + picCode);
                if (!picCode.startsWith("$")){
                    continue;
                }
                picCode = picCode.substring(picCode.indexOf("{") + 1, picCode.indexOf("}"));
                // 获取文档中的编码时，SZNF可以识别出SZNF-1,SZNF
                String code_ = null;
                if (picCode.contains("-")) {
                    String[] s1 = picCode.split("-");
                    if (!ObjectUtils.isEmpty(s1)) {
                        code_ = s1[0];
                    }
                }
                if (codeList.contains(code_)){
                    codes.add(code_);
                    continue;
                }
                if (!codeList.contains(picCode)){
                    continue;
                }
                log.info("【获取图片编码】:" + picCode);
                codes.add(picCode);
            }
        }
    }


    /**
     * 处理接受的list实体
     * @param obj
     * @return
     */
    private static List<List> handleTables(Object obj) {
        log.info("【处理接受的list实体】 开始" + obj);
        List<List> allList = new ArrayList<>();
        if (obj == null) {
            return allList;
        }
        if (obj instanceof ArrayList<?>) {
            List obj1 = (List)obj;
            for (Object o : obj1) {
                List<String> rowList = new ArrayList<>();
                if (o instanceof ArrayList<?>){
                    return (List)obj;
                }
                Class<?> aClass = o.getClass();
                if (aClass.getName().equals("com.alibaba.fastjson.JSONArray")){
                    List<Object> list = JSONObject.parseArray(o.toString(), Object.class);
                    for (Object o1 : list) {
                        rowList.add(o1.toString());
                    }
                    allList.add(rowList);
                    continue;
                }
                Field[] fields = aClass.getDeclaredFields();
                try {
                    for (Field field : fields) {
                        field.setAccessible(true);
                        rowList.add( field.get(o).toString());
                    }
                    allList.add(rowList);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }
        log.info("【处理接受的list实体】 结束" + allList);
        return allList;
    }

    private static int getPictureType(String picType){
        int res = CustomXWPFDocument.PICTURE_TYPE_PICT;
        if(picType != null){
            if(picType.equalsIgnoreCase("png")){
                res = CustomXWPFDocument.PICTURE_TYPE_PNG;
            }else if(picType.equalsIgnoreCase("dib")){
                res = CustomXWPFDocument.PICTURE_TYPE_DIB;
            }else if(picType.equalsIgnoreCase("emf")){
                res = CustomXWPFDocument.PICTURE_TYPE_EMF;
            }else if(picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")){
                res = CustomXWPFDocument.PICTURE_TYPE_JPEG;
            }else if(picType.equalsIgnoreCase("wmf")){
                res = CustomXWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

}
