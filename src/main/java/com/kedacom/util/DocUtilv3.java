package com.kedacom.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.aspose.words.SaveFormat;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;
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
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 15:21
 * @apiNote : docUtils第三版
 */
@Slf4j
public class DocUtilv3 {
    private static final String GXK_FLAG_TRUE = "GXK_FLAG_TRUE_CONST";
    private static final String GXK_FLAG_FALSE = "GXK_FLAG_FALSE_CONST";
    private static final String WINGDINGS_SQUARE = "Wingdings 2";
    private static final String WINGDINGS_SQUARE_TURE = "\u0052";
    private static final String WINGDINGS_SQUARE_FALSE = "\u0030";
    private static final String FS_PIC_PREFIX = "/var/ftphome/kvms3000/fs/openImage/";
    private static final String FS_NET_PIC_PREFIX = "/var/ftphome/kvms3000http";
    private static final String FS_PIC_LOCAL_PREFIX = "/var/ftphome/kvms3000/files/uploads/images/";
    private static final String DOCX_TYPE = "docx";
    private static final String PDF_TYPE_POSTFIX = ".pdf";
    private static final String HTML_TYPE_POSTFIX = ".html";
    private static final String FILE_POINT = ".";
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
    // 只获取后缀名
    private static String getFileExtension (String name) {
        return name.substring(name.lastIndexOf(FILE_POINT) + 1);
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
        String tableReg="<w:tbl[Dd]escription\\sw:val=\"(?<code>[A-Z]+)\"/>";
        Pattern compile = Pattern.compile(tableReg);
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        String tblpr = tblPr.toString();
        if (!ObjectUtils.isEmpty(tblPr)) {
            Matcher matcher = compile.matcher(tblpr);
            while (matcher.find()){
                String codeStr = matcher.group("code");
                for (String code : codeList) {
                    if (codeStr.equals(code)) {
                        log.info("【获取表格编码】:{}",codeStr);
                        codes.add(code);
                    }
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
                com.aspose.words.Document doc = new com.aspose.words.Document(filePath);
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
    // 获取文件名
    private static String getFileName (String name) {
        return name.substring(0, name.lastIndexOf(FILE_POINT));
    }
    /**
     * 打开pdf文档
     * @param fileLocalName pdf文档全路径
     * @param response
     */
    public static void openPdf(String fileLocalName, HttpServletResponse response) {
        getFileResponse(fileLocalName, MediaType.APPLICATION_PDF_VALUE,response);
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
    private static void picCodeHandler(Map<String, Object> param, List<XWPFRun> runs, XWPFParagraph paragraph, CustomXWPFDocument document) throws FileNotFoundException, InvalidFormatException {

        List<XWPFRun> runList = paragraph.getRuns();
        for (int i = 0; i < runList.size(); i++) {
            List<XWPFPicture> pictures = runList.get(i).getEmbeddedPictures();
            for (XWPFPicture picture : pictures) {
                String desc = picture.getDescription();
                if (desc.startsWith("$")) {
                    desc = desc.substring(desc.indexOf("{") + 1, desc.indexOf("}"));
                    if (param.keySet().contains(desc)) {
                        log.debug("【图片编码】:" + desc);
                        CTPicture ctPicture = picture.getCTPicture();
                        String picPaths = (String) param.get(desc);
                        //图片编码对应的值为空时不进行处理
                        if (StringUtils.isEmpty(picPaths)) {
                            continue;
                        }
                        replaceImageInParagraph(picPaths, ctPicture, paragraph, document, i, desc);
                    }
                }
            }
        }
    }
    private static void replaceMultipleImageInParagraph(String[] split1, CTPicture ctPicture, XWPFParagraph paragraph,
                                                        CustomXWPFDocument document, int i, String desc)
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
            if (StringUtils.isEmpty(blipId)) {
                continue;
            }
            document.createPicture(blipId, (int) id, cx, cy, desc, paragraph, i);
        }

    }
    private static String handleBlipId(String picPath, CustomXWPFDocument document) throws FileNotFoundException, InvalidFormatException {
        //判断picPath路径  /var/ftphome/kvms3000/fs/openImage/20220728095420850_0.jpg
        //拼接成           /var/ftphome/kvms3000/files/uploads/images/20220728/20220728095420850_0.jpg
        //拼接成           /var/ftphome/kvms3000http://172.16.231.169:19505/fs/openImage/20220817090913763_0.JPG
        String[] split = picPath.split("\\.");
        String picType = split[split.length - 1];
        File file = new File(picPath);
        if (picPath.startsWith(FS_PIC_PREFIX)) {
            String[] split2 = picPath.split("/");
            String picName = split2[split2.length - 1];
            String subPicPath = picName.substring(0, 8);
            picPath = FS_PIC_LOCAL_PREFIX + subPicPath + File.separator + picName;
        }
        if (picPath.startsWith(FS_NET_PIC_PREFIX)) {
            String picNetUrl = picPath.substring(21);
            if (!verifyImage(getPicInputStream(picNetUrl))) {
                log.error("获取网络地址流信息失败");
                return null;
            }
            InputStream picInputStream = getPicInputStream(picNetUrl);
            if (null != picInputStream) {
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
    private static InputStream getPicInputStream(String url) {
        try {
            URL httpUrl = new URL(url);
            HttpURLConnection conn = (HttpURLConnection) httpUrl.openConnection();
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
    private static int getPictureType(String picType) {
        int res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") || picType.equalsIgnoreCase("jpeg")) {
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }
    private static void replaceImageInParagraph(String picPaths, CTPicture ctPicture, XWPFParagraph paragraph,
                                                CustomXWPFDocument document, int i, String desc)
            throws FileNotFoundException, InvalidFormatException {
        //多张图片处理
        String[] split1 = picPaths.split(";");
        if (split1.length == 1) {
            for (String picPath : split1) {
                String blipId = handleBlipId(picPath, document);
                if (StringUtils.isEmpty(blipId)) {
                    continue;
                }
                ctPicture.getBlipFill().getBlip().setEmbed(blipId);
            }
        } else {
            replaceMultipleImageInParagraph(split1, ctPicture, paragraph, document, i, desc);
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
        if (obj instanceof List) {
            List obj1 = (List)obj;
            for (Object o : obj1) {
                if (o instanceof List){
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
    private static void insertTableByCode(List<List<Object>> lists , String string , XWPFTable table){
        if (!CollectionUtils.isEmpty(lists)){
            if(string.equals("SSCWJCJL")){
                //处理登记表中的随身财物检查记录
                //insertTable(table, lists, 5 , 11 , 1,1);
                insertRowAndCopyStyle(10, 15, lists, table, "仿宋", 12);
            } else  if(string.equals("HDJL")){
                //处理登记表中的活动记录
                //insertTable(table, lists, 11 , 2 , 1,1);
                insertRowAndCopyStyle(1, 11, lists, table, "仿宋", 12);
            }else {
                // 处理清单中的字体
                //insertTable(table, lists, table.getNumberOfRows() - 2 , 2 , 1,0);
                //startRowIndex是表格下标,下标从0开始,这里的1是指起始行从第二行包括第二行,容量为14,1+14=15,endRowIndex为15,不包括下标为15的行
                insertRowAndCopyStyle(1, 15, lists, table, "仿宋_GB2312", 14);
            }
        }
    }

    /**
     * 循环插入表格数据
     *  @param startRowIndex 表格填充的起始下标位置,包括该下标
     * @param endRowIndex   表格的结束行,不包括该下标
     * @param tableList     要插入的表格数据
     * @param table         需要插入的表格,如果为空则抛出异常
     * @param fontFamily    字体
     * @param fontSize      字体大小
     */
    private static void insertRowAndCopyStyle(int startRowIndex,
                                              int endRowIndex,
                                              List<List<Object>> tableList,
                                              XWPFTable table,
                                              String fontFamily,
                                              Integer fontSize) {
        XWPFTableRow sourceRow = table.getRow(startRowIndex);
        if (sourceRow==null){
            throw new IllegalArgumentException("startRowIndex参数错误,获取不到表格行");
        }
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




    public static String replaceWordCode(Map<String, Object> param, String srcWordPath) {
        checkParam(param, srcWordPath);
        String targetPath = getTargetPath(srcWordPath);
        try (FileOutputStream fos = new FileOutputStream(targetPath);
             CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(srcWordPath))) {
            //获取所有段落
            List<XWPFParagraph> paragraphs = document.getParagraphs();
            if (!CollectionUtils.isEmpty(paragraphs)) {
                for (XWPFParagraph paragraph : paragraphs) {
                    paragraphHandler(paragraph, param, document);
                }
            }
            creatTableInParagraph(document,param);
            //处理表格 还有表格增行操作
            List<XWPFTable> tables = document.getTables();
            for (XWPFTable table : tables) {
                List<XWPFTableRow> rows = table.getRows();
                for (XWPFTableRow row : rows) {
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    for (XWPFTableCell tableCell : tableCells) {
                        List<XWPFParagraph> tableCellParagraphs = tableCell.getParagraphs();
                        for (XWPFParagraph paragraph : tableCellParagraphs) {
                            paragraphHandler(paragraph, param, document);
                        }
                    }
                }
            }
            //页眉
            List<XWPFHeader> headerList = document.getHeaderList();
            for (XWPFHeader xwpfHeader : headerList) {
                List<XWPFParagraph> xwpfHeaderParagraphs = xwpfHeader.getParagraphs();
                for (XWPFParagraph paragraph : xwpfHeaderParagraphs) {
                    paragraphHandler(paragraph, param, document);
                }
            }
            //页脚
            List<XWPFFooter> footerList = document.getFooterList();
            for (XWPFFooter xwpfFooter : footerList) {
                List<XWPFParagraph> xwpfFooterParagraphs = xwpfFooter.getParagraphs();
                for (XWPFParagraph paragraph : xwpfFooterParagraphs) {
                    paragraphHandler(paragraph, param, document);
                }
            }
            document.write(fos);
        } catch (Exception e) {
            log.info("替换模板异常",e);
        }
        return targetPath;
    }

    private static void paragraphHandler(XWPFParagraph paragraph, Map<String, Object> param, CustomXWPFDocument document) throws FileNotFoundException, InvalidFormatException {
        List<XWPFRun> runs = paragraph.getRuns();
        //跨行编码文本预处理,获取文本缓存
        List<String> idForRuns = runsMultilineCodeHandler(paragraph, runs);
        //文本勾选框编码处理
        getGXKParam(param, idForRuns);
        //ITEM系列编码处理
        getITEMParam(param,idForRuns);
        //文本编码处理
        textCodeHandler(param, runs, idForRuns);
        //勾选框处理
        gxkCodeHandler(paragraph, runs, idForRuns);
        //图片编码处理
        picCodeHandler(param, runs, paragraph, document);
    }

    private static Map<String, Object> handItemCodes(Map<String,Object> param){
        Object item = param.get("ITEM");
        Map<String, Object> itemMap = new HashMap<>();
        if (!ObjectUtils.isEmpty(item)) {
            for (Map.Entry<String, Object> entry : param.entrySet()) {
                itemMap.put(entry.getKey() + "-" + item , entry.getValue());
            }
        }
        return itemMap;
    }
    private static void getITEMParam(Map<String, Object> param, List<String> idForRuns) {
        Object item = param.get("ITEM");
        if (item !=null){
            Map<String,String> itemMap=new HashMap<>();
            //匹配所有item编码
            String reg="[\\u4e00-\\u9fa5_a-zA-Z0-9]+-[0-9]+";
            Pattern codeCompile = Pattern.compile(reg);
            String dollarReg="\\$\\{[\\u4e00-\\u9fa5_a-zA-Z0-9]+-[0-9]+}";
            Pattern dollarCompile = Pattern.compile(dollarReg);
            for (String idForRun : idForRuns) {
                Matcher matcher = dollarCompile.matcher(idForRun);
                while (matcher.find()){
                    String dollarCode = matcher.group();
                    Matcher codeMatcher = codeCompile.matcher(dollarCode);
                    if (codeMatcher.find()) {
                        String code_num = codeMatcher.group();
                        if (itemMap.get(code_num)==null){
                            String[] split = code_num.split("-");
                            String itemNumber = split[split.length - 1];
                            if (Objects.equals(itemNumber, item.toString())) {
                                Object o = param.get(split[0]);
                                if (o!=null) {
                                    itemMap.put(code_num,o.toString());
                                }else {
                                    itemMap.put(code_num,"    ");
                                }
                            }
                            //其他item默认为空
                            else {
                                itemMap.put(code_num,"    ");
                            }
                        }
                    }
                }
            }
            param.putAll(itemMap);
        }
    }


    private static List<String> runsMultilineCodeHandler(XWPFParagraph paragraph, List<XWPFRun> runs) {
        List<String> idForRuns = new ArrayList<>(256);
        //文本缓存,与id对应
        int startRunIndex = -1;
        int endRunIndex = -1;
        StringBuilder startRunText = new StringBuilder();
        StringBuilder endRunText = new StringBuilder();
        for (int i = 0; i < runs.size(); i++) {
            //获取本段run的文本
            String text = runs.get(i).toString();
            idForRuns.add(text);
            //匹配残缺的编码标识
            String endRunMutilatedReg = "\\$\\{?[^}]*$";
            Pattern compile = Pattern.compile(endRunMutilatedReg);
            Matcher matcher = compile.matcher(text);
            //寻找下一段匹配的编码末尾
            if (startRunIndex != -1) {
                String endRunReg = "^\\S*}";
                Pattern pattern = Pattern.compile(endRunReg);
                Matcher matcher1 = pattern.matcher(text);
                if (matcher1.find()) {
                    //记录末尾位置与字符下标
                    endRunIndex = i;
                    int end = matcher1.end();
                    //编码处理
                    //}之后的文本
                    String substrAfter = text.substring(end);
                    endRunText.append(substrAfter);
                    //}
                    String substrBefore = text.replace(substrAfter, "");
                    startRunText.append(substrBefore);
                }else {
                    //没到末尾,只附加文本
                    startRunText.append(text);
                }
            }
            if (startRunIndex == -1&&matcher.find()) {
                //记录残缺编码起始位置
                startRunIndex = i;
                startRunText.append(text);
            }
            //删除中间跨行文本并替换开头结尾的runs文本
            if (startRunIndex != -1 && endRunIndex != -1) {
                //处理文本
                XWPFRun startRun = runs.get(startRunIndex);
                XWPFRun endRun = runs.get(endRunIndex);
                for (int index = startRunIndex; index <= endRunIndex; index++) {
                    if (index > startRunIndex && index < endRunIndex) {
                        //删除后自动补位,对应下标相应-1
                        paragraph.removeRun(index);
                        endRunIndex--;
                        idForRuns.remove(index);
                        index--;
                        i--;
                    }
                }
                //替换内容
                //pos代表w:t标签的下标
                startRun.setText(startRunText.toString(), 0);
                idForRuns.set(startRunIndex, startRunText.toString());
                endRun.setText(endRunText.toString(), 0);
                //idForRuns.set(endRunIndex, endRunText.toString());
                //删除末尾的run,计数器减一,重新获取run校验以防出现两个残缺编码在同一列
                idForRuns.remove(endRunIndex);
                i--;
                //重置标志位
                startRunIndex = -1;
                endRunIndex = -1;
                startRunText.delete(0, startRunText.length());
                endRunText.delete(0, endRunText.length());
            }
        }
        return idForRuns;
    }

    private static void textCodeHandler(Map<String, Object> param, List<XWPFRun> runs, List<String> idForRuns) {
        for (int i = 0; i < runs.size(); i++) {
            //获取本段run的文本
            String text = idForRuns.get(i);
            String newText = text;
            String reg = "\\$\\{\\w+[^$]+}";
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(text);
            //匹配文本key
            while (matcher.find()) {
                for (Map.Entry<String, Object> entry : param.entrySet()) {
                    if (Objects.equals(entry.getKey(), "ITEM")){
                        newText=newText.replace("${"+entry.getKey()+"}","");
                    }else {
                        newText = newText.replace("${"+entry.getKey()+"}", entry.getValue().toString());
                    }
                }
            }
            //如果文本有编码被替换,则更新run
            if (!text.equals(newText)) {
                XWPFRun xwpfRun = runs.get(i);
                //todo 更新run格式
                //xwpfRun.setFontFamily();
                //xwpfRun.setFontSize();
                //xwpfRun.setUnderline();
                //xwpfRun.setBold();
                xwpfRun.setText(newText, 0);
                idForRuns.set(i, newText);
            }
        }
    }

    private static void gxkCodeHandler(XWPFParagraph paragraph, List<XWPFRun> runs, List<String> idForRuns) {
        for (int i = 0; i < runs.size(); i++) {
            String nowRuns = idForRuns.get(i);
            String reg=GXK_FLAG_TRUE+"|"+GXK_FLAG_FALSE;
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(nowRuns);
            if (matcher.find()){
                String substring = nowRuns.substring(matcher.start(), matcher.end());
                //这一行仅有这个文本
                if (substring.length()==nowRuns.length()){
                    paragraph.removeRun(i);
                    XWPFRun gxkRun = paragraph.insertNewRun(i);
                    if (Objects.equals(matcher.group(), GXK_FLAG_FALSE)) {
                        gxkRun.setText(WINGDINGS_SQUARE_FALSE);
                    }else {
                        gxkRun.setText(WINGDINGS_SQUARE_TURE);
                    }
                    gxkRun.setFontFamily(WINGDINGS_SQUARE);
                    idForRuns.set(i, "勾选框替换占位符");
                }else {
                    String beforeText = nowRuns.substring(0, matcher.start());
                    String afterText = nowRuns.substring(matcher.end());
                    //提前读取格式
                    XWPFRun afterRuns = runs.get(i);
                    UnderlinePatterns underline = afterRuns.getUnderline();
                    boolean bold = afterRuns.isBold();
                    afterRuns.setText(afterText, 0);
                    idForRuns.set(i, afterText);
                    //删除替换为勾选框
                    XWPFRun gxkRun = paragraph.insertNewRun(i);
                    if (Objects.equals(matcher.group(), GXK_FLAG_FALSE)) {
                        gxkRun.setText(WINGDINGS_SQUARE_FALSE);
                    }else {
                        gxkRun.setText(WINGDINGS_SQUARE_TURE);
                    }
                    gxkRun.setFontFamily(WINGDINGS_SQUARE);
                    idForRuns.add(i, "勾选框替换占位符");
                    if (!StringUtils.isEmpty(beforeText)){
                        XWPFRun beforeRun = paragraph.insertNewRun(i);
                        beforeRun.setUnderline(underline);
                        beforeRun.setBold(bold);
                        beforeRun.setText(beforeText);
                        idForRuns.add(i, beforeText);
                    }
                }
            }
        }
    }

    private static void getGXKParam(Map<String, Object> param, List<String> idForRuns) {
        //查找每行编码是勾选框的,结合param参数,将选中的标记为true,未选中的标记为false
        Map<String, String> dollarParamForFlag = new HashMap<>();
        String gxkReg = "[A-Z0-9]+_[^}-]+";
        String dollarGxkReg = "\\$\\{[A-Z0-9]+_[^}-]+}";
        Pattern gxkCompile = Pattern.compile(gxkReg);
        Pattern dollarCompile = Pattern.compile(dollarGxkReg);
        for (String idForRun : idForRuns) {
            Matcher dollarMatcher = dollarCompile.matcher(idForRun);
            //匹配每一个完整的${XXX_XXX}编码
            while (dollarMatcher.find()) {
                String dollarGxkStr = dollarMatcher.group();
                log.info("文本中有勾选框编码:{}", dollarGxkStr);
                Matcher gxkMatcher = gxkCompile.matcher(dollarGxkStr);
                //匹配${XXX_XXX}内的XXX_XXX
                if (gxkMatcher.find()) {
                    String gxkCode = gxkMatcher.group();
                    String[] split = gxkCode.split("_");
                    String code = split[0];
                    String suffix = split[1];
                    log.info("获取编码:{}", code);
                    Object o = param.get(code);
                    boolean flag=false;
                    //如果没有参数,则代表该勾选框没有被选中
                    if (o!=null){
                        String values = o.toString();
                        log.info("param编码值为:{}", values);
                        //这里的编码可能是XXX,XXX,XXX的多选框格式
                        String[] valueList = values.split(",");
                        for (String value : valueList) {
                            if (Objects.equals(suffix, value)) {
                                flag=true;
                                break;
                            }
                        }
                    }else {
                        log.info("参数为空,勾选框值为false");
                    }
                    if (flag){
                        log.info("勾选框值为true");
                        dollarParamForFlag.put(gxkCode, GXK_FLAG_TRUE);
                    }else {
                        log.info("勾选框值为false");
                        dollarParamForFlag.put(gxkCode, GXK_FLAG_FALSE);
                    }
                }
            }
        }
        param.putAll(dollarParamForFlag);
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