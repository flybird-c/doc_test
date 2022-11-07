package com.kedacom.controller;


import com.kedacom.exception.ServiceException;
import com.kedacom.util.DocUtilv3;
import com.kedacom.util.PdfToImageUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.web.bind.annotation.*;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.List;
import java.util.Map;

@Api(tags = "测试文书工具类API")
@RestController
public class TestDocUtilsController {

    @PutMapping(value = "/test/codes")
    @ApiOperation(value = "获取word文档中的编码")
    public List getWordCode(@RequestBody List<String> codeList , String srcWordPath) throws ServiceException {
        return DocUtilv3.getDocCodes(codeList,srcWordPath);
    }

    @PutMapping(value = "/test/code")
    @ApiOperation(value = "替换word文档中的编码")
    public String getWordCode(@RequestBody Map<String,Object> map , String srcWordPath ) throws ServiceException, IOException {
        return DocUtilv3.replaceWordCode(map,srcWordPath);
    }

    @GetMapping(value = "/test/pdf")
    @ApiOperation("根据word文件的路径获取pdf路径")
    public String getWordByRelativeFilePath( @RequestParam(value = "filePath") String filePath
    ) throws ServiceException {
        return DocUtilv3.getPdfByWordPath(filePath);
    }

    @GetMapping("/test/openPdf")
    @ApiOperation(value = "打开PDF")
    public void openPdf(@RequestParam("fileLocalName") String fileLocalName, HttpServletResponse response) throws ServiceException {
        DocUtilv3.openPdf(fileLocalName,response);
    }
    @GetMapping(value = "/test/html")
    @ApiOperation("根据word文件的路径获取html路径")
    public String getHtmlByRelativeFilePath( @RequestParam(value = "filePath") String filePath
    ) throws ServiceException {
        return DocUtilv3.getHtmlByWordPath(filePath);
    }

    @GetMapping("/test/openHtml")
    @ApiOperation(value = "打开HTML")
    public void openHtml(@RequestParam("fileLocalName") String fileLocalName, HttpServletResponse response) throws ServiceException {
        DocUtilv3.openHtml(fileLocalName,response);
    }

    @GetMapping("/test/pdfToPic")
    @ApiOperation(value = "pdf转图片")
    public List<String> pdfToPic(String pdfPath) throws ServiceException {
        return PdfToImageUtil.pdfToImages(pdfPath);
    }
}
