package com.kedacom.util;



import org.apache.xmlbeans.XmlToken;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;

import java.io.IOException;
import java.io.InputStream;

public class CustomXWPFDocument extends XWPFDocument {
    public CustomXWPFDocument(InputStream in) throws IOException {
        super(in);
    }

    public CustomXWPFDocument(OPCPackage pkg) throws IOException {
        super(pkg);
    }

    public void createPicture(String blipId, int id, long width, long height, String desc, XWPFParagraph paragraph,int i) {

        //给段落插入图片
        CTDrawing ctDrawing = paragraph.insertNewRun(i).getCTR().addNewDrawing();
        CTInline inline = ctDrawing.addNewInline();
        if (desc.equals("DZQZKEDACOMTEST")) {
//            CTAnchor ctAnchor = ctDrawing.addNewAnchor();
//            CTPosH ctPosH = ctAnchor.addNewPositionH();
//            ctPosH.setPosOffset(3450590);
//            CTPosV ctPosV = ctAnchor.addNewPositionV();
//            ctPosV.setPosOffset(12065);
        }

//        CTDrawing drawing = paragraph.createRun().getCTR().addNewDrawing();
//        CTAnchor anchor = drawing.addNewAnchor();
//        anchor.setBehindDoc(false);
//        anchor.setLocked(false);
//        anchor.setLayoutInCell(true);
//        anchor.setAllowOverlap(true);
//        CTAnchor inline = anchor;

        String picXml = "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                "         <pic:nvPicPr>" +
                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
//                "              <a:picLocks noChangeAspect=\"1\" noChangeArrowheads=\"1\"/>" +
                "            <pic:cNvPicPr/>" +
                "         </pic:nvPicPr>" +
                "         <pic:blipFill>" +
                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                "            <a:stretch>" +
                "               <a:fillRect/>" +
                "            </a:stretch>" +
                "         </pic:blipFill>" +
                "         <pic:spPr>" +
                "            <a:xfrm>" +
                "               <a:off x=\"0\" y=\"0\"/>" +
                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +
                "            </a:xfrm>" +
                "            <a:prstGeom prst=\"rect\">" +
                "               <a:avLst/>" +
                "            </a:prstGeom>" +
                "         </pic:spPr>" +
                "      </pic:pic>" +
                "   </a:graphicData>" +
                "</a:graphic>";

        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();

        XmlToken xmlToken = null;

        try {
            xmlToken = xmlToken = XmlToken.Factory.parse(picXml);

        } catch(XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        //graphicData.set(xmlToken);

        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);

        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(id);
        docPr.setName("Picture " + id);
        docPr.setDescr(desc);
    }
}


