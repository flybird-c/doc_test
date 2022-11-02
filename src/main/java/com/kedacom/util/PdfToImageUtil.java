package com.kedacom.util;

import com.lowagie.text.pdf.PdfReader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class PdfToImageUtil {

    public static List<String> pdfToImages(String pdfPath) {
        PdfToImageUtil pdfToImageUtil = new PdfToImageUtil();
        List<File> files = pdfToImageUtil.pdfToImage(pdfPath);
        List<String> pdfToPicPaths = new ArrayList<>();
        for (File file : files) {
            String AbsolutePath=file.getAbsolutePath();
            pdfToPicPaths.add(AbsolutePath);
        }
        return pdfToPicPaths;
    }

    private List<File> pdfToImage(String PdfFilePath) {
        File file = new File(PdfFilePath);
        List<File> fileList = new ArrayList<File>();
        @SuppressWarnings("resource")//抑制警告
                PDDocument pdDocument = new PDDocument();
        String[] split = PdfFilePath.split("\\.");
        String filePdfPath = split[0];
        String[] split1 = filePdfPath.split("\\/");
        String fileName = split1[split1.length - 1];
        try {
            String imgFolderPath = filePdfPath + File.separator;
            if (createDirectory(imgFolderPath)) {
                pdDocument = PDDocument.load(file);
                PDFRenderer renderer = new PDFRenderer(pdDocument);
                PdfReader reader = new PdfReader(PdfFilePath);
                StringBuffer imgFilePath = null;
                for (int i = 0; i < reader.getNumberOfPages(); i++) {
                    String imgFilePathPrefix = imgFolderPath + File.separator + fileName;
                    imgFilePath = new StringBuffer();
                    imgFilePath.append(imgFilePathPrefix);
                    imgFilePath.append("-");
                    imgFilePath.append(String.valueOf(i));
                    imgFilePath.append(".jpg");
                    File dstFile = new File(imgFilePath.toString());
                    BufferedImage image = renderer.renderImageWithDPI(i, 150);
                    ImageIO.write(image, "png", dstFile);
                    fileList.add(dstFile);
                }
                return fileList;
            } else {
                return null;
            }
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }

    private boolean createDirectory(String folder) {
        File dir = new File(folder);
        if (dir.exists()) {
            return true;
        } else {
            return dir.mkdirs();
        }
    }

    private void delFolder(String folderPath) {
        try {
            delAllFile(folderPath);
            String filePath = folderPath;
            filePath = filePath.toString();
            File myFilePath = new File(filePath);
            myFilePath.delete();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    private boolean delAllFile(String path) {
        boolean flag = false;
        File file = new File(path);
        if (!file.exists()) {
            return flag;
        }
        if (!file.isDirectory()) {
            return flag;
        }
        String[] tempList = file.list();
        File temp = null;
        for (int i = 0; i < tempList.length; i++) {
            if (path.endsWith(File.separator)) {
                temp = new File(path + tempList[i]);
            } else {
                temp = new File(path + File.separator + tempList[i]);
            }
            if (temp.isFile()) {
                temp.delete();

            }
            if (temp.isDirectory()) {
                delAllFile(path + "/" + tempList[i]);
                delFolder(path + "/" + tempList[i]);
                flag = true;
            }
        }
        return flag;
    }

}

