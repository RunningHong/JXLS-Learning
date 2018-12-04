package com.zh.jxls.util;

import org.apache.poi.hssf.converter.ExcelToHtmlConverter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.FileUtils;

import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

/**
 *
 */
public class ExcelToHtmlUtil {

    /**
     * 将excel转换为html
     * @Author RunningHong
     * @Date 2018/12/3 16:59
     * @Param
     * @return
     */
    public static void excelToHtml(InputStream input, OutputStream output) throws Exception {
        HSSFWorkbook workbook;
        try {
            workbook = new HSSFWorkbook(input);
        } catch (Exception exc) {
            exc.printStackTrace();
            return;
        }

        ExcelToHtmlConverter excelToHtmlConverter = new ExcelToHtmlConverter(XMLHelper.getDocumentBuilderFactory()
                                                                            .newDocumentBuilder().newDocument());
        // 去掉头
        excelToHtmlConverter.setOutputColumnHeaders(false);

        // 去掉行号
        excelToHtmlConverter.setOutputRowNumbers(false);

        // excelToHtmlConverter.setOutputSheetHeaders(false);
        excelToHtmlConverter.processWorkbook(workbook);

        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        transformer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
        transformer.setOutputProperty(OutputKeys.INDENT, "no");
        transformer.setOutputProperty(OutputKeys.METHOD, "html");

        transformer.transform(new DOMSource(excelToHtmlConverter.getDocument()),
                              new StreamResult(output));



    }

}
