package com.kedacom.util;

import com.aspose.words.SaveFormat;
import lombok.SneakyThrows;
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
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

/**
 * @author : lzp
 * @version 1.0
 * @date : 2022/10/31 15:21
 * @apiNote : TODO
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
                    paragraphHandler(paragraph, param, document);
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
        //文本编码处理
        textCodeHandler(param, runs, idForRuns);
        //勾选框处理
        gxkCodeHandler(paragraph, runs, idForRuns);
        //图片编码处理
        picCodeHandler(param, runs, paragraph, document);
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
                    newText = newText.replace(entry.getKey(), entry.getValue().toString());
                }
            }
            //如果文本有编码被替换,则更新run
            if (!text.equals(newText)) {
                XWPFRun xwpfRun = runs.get(i);
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
                    if (!StringUtils.isEmpty(afterText)){
                        afterRuns.setText(afterText, 0);
                        idForRuns.set(i, afterText);
                    }
                    //删除替换为勾选框
                    paragraph.removeRun(i);
                    XWPFRun gxkRun = paragraph.insertNewRun(i);
                    if (Objects.equals(matcher.group(), GXK_FLAG_FALSE)) {
                        gxkRun.setText(WINGDINGS_SQUARE_FALSE);
                    }else {
                        gxkRun.setText(WINGDINGS_SQUARE_TURE);
                    }
                    gxkRun.setFontFamily(WINGDINGS_SQUARE);
                    idForRuns.set(i, "勾选框替换占位符");
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
        List<String> removeList = new ArrayList<>();
        String gxkReg = "\\w+_[^}]+";
        String dollarGxkReg = "\\$\\{\\w+_[^$]+}";
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
                    log.info("获取编码:{}", split[0]);
                    Object o = param.get(split[0]);
                    String values = o.toString();
                    log.info("param编码值为:{}", values);
                    String[] valueList = values.split(",");
                    boolean flag = false;
                    for (String value : valueList) {
                        if (Objects.equals(split[1], value)) {
                            flag = true;
                            break;
                        }
                    }
                    //如果值相等,则增加true标记,不相等增加false标记
                    if (flag) {
                        log.info("勾选框值为true");
                        dollarParamForFlag.put(dollarGxkStr, GXK_FLAG_TRUE);
                    } else {
                        log.info("勾选框值为false");
                        dollarParamForFlag.put(dollarGxkStr, GXK_FLAG_FALSE);
                    }
                    //原本param的参数移除替换
                    removeList.add(split[0]);
                }
            }
        }
        removeList.forEach(param::remove);
        param.putAll(dollarParamForFlag);
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
