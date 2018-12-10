package com.zh.jxls.util;

import com.artofsolving.jodconverter.*;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.AbstractOpenOfficeDocumentConverter;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import org.apache.commons.io.FilenameUtils;
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
 * Excel转换为html工具类
 * @Author RunningHong
 * @Date 2018/12/8 15:32
 */
public class ExcelToHtmlUtil {

    static Logger log = LoggerFactory.getLogger(ExcelToHtmlUtil.class);

    /**
     * 使用openOffice将excel流转换为HTML流并放于os中
     * @Author RunningHong
     * @Date 2018/12/7 17:56
     * @Param
     * @return
     */
    public static void excelStreamToHtmlStreamByOpenOffice(InputStream is, OutputStream os) {
        // 打开OpenOffice服务
        OpenOfficeUtil.startOpenOffice();
        OpenOfficeConnection connection = new SocketOpenOfficeConnection(8100);
        try {
            connection.connect();
        } catch (Exception e) {
            System.err.println("OpenOffice文件转换出错，请检查OpenOffice服务是否启动。");
        }

        // 转换器
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);

        // html格式
        DocumentFormat html = new DocumentFormat("HTML", DocumentFamily.TEXT, "text/html", "html");
        html.setExportFilter(DocumentFamily.PRESENTATION, "impress_html_Export");
        html.setExportFilter(DocumentFamily.SPREADSHEET, "HTML (StarCalc)");
        html.setExportFilter(DocumentFamily.TEXT, "HTML (StarWriter)");

        // xls格式
        DocumentFormat xls = new DocumentFormat("Microsoft Excel", DocumentFamily.SPREADSHEET, "application/vnd.ms-excel", "xls");
        xls.setExportFilter(DocumentFamily.SPREADSHEET, "MS Excel 97");

        // 将xls流装换为html流
        converter.convert(is, xls, os, html);
        connection.disconnect();

        log.info("根据Excel模板使用OpenOffice预览HTML成功！");
    }


    /**
     * 使用poi将Excel文件流out转化为HTML，数据存储在os中
     * 使用poi
     * @Author RunningHong
     * @Date 2018/12/3 16:59
     * @Param
     * @return
     */
    public static void excelStreamToHtmlStreamByPoi(InputStream is, OutputStream os) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(is);

        // 修改了ExcelToHtmlConverter的一点内容
        ExcelToHtmlConverterModify excelToHtmlConverter = new ExcelToHtmlConverterModify(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        // ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());


        // 去掉列号(Excel的A B C D列号)
        excelToHtmlConverter.setOutputColumnHeaders(false);
        // 去掉行号(Excel的1 2 3 4行号)
        excelToHtmlConverter.setOutputRowNumbers(false);
        excelToHtmlConverter.setUseDivsToSpan(false);
        excelToHtmlConverter.setOutputHiddenColumns(false);
        excelToHtmlConverter.setOutputHiddenRows(false);
        excelToHtmlConverter.setOutputLeadingSpacesAsNonBreaking(false);

        excelToHtmlConverter.processWorkbook(workbook);

        // 转化为HTML
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "no");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");
        transformer.transform(new DOMSource(excelToHtmlConverter.getDocument()), new StreamResult(os));

        log.info("根据Excel模板使用POI预览HTML成功！");
    }

}
