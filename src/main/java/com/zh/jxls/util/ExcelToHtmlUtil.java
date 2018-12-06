package com.zh.jxls.util;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.List;

/**
 *
 */
public class ExcelToHtmlUtil {

    static Logger log = LoggerFactory.getLogger(ExcelToHtmlUtil.class);

    /**
     * 将Excel文件流out转化为HTML，数据存储在response中
     * @Author RunningHong
     * @Date 2018/12/3 16:59
     * @Param
     * @return
     */
    public static void excelToHtml(InputStream input, OutputStream output) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(input);

        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        // 去掉头
        excelToHtmlConverter.setOutputColumnHeaders(false);
        // 去掉行号
        excelToHtmlConverter.setOutputRowNumbers(false);
        excelToHtmlConverter.processWorkbook(workbook);

        // 读取Excel文件的图片到指定位置
        String picSavePath = "D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\out\\generatePic\\";
        ExcelUtil.generateExcelPicToFile(workbook, picSavePath);

        // 转化为HTML
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "no");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        transformer.transform(new DOMSource(excelToHtmlConverter.getDocument()), new StreamResult(output));

        log.info("根据Excel模板预览HTML成功！");


        /*// 将生成的图片Excel写出来
        String picSaveExcelName = picSavePath + "图片Excel.xls";
        FileOutputStream picFileOut = new FileOutputStream(picSaveExcelName);
        workbook.write(picFileOut);*/
    }

}
