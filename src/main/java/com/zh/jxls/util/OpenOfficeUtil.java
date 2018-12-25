package com.zh.jxls.util;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFamily;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.InputStream;
import java.io.OutputStream;

/**
 * 需要在配置文件中配置OpenOffice的相关参数
 * OpenOfficeUtil
 * @author RunningHong
 * @since 2018-12-07 17:00
 */
public class OpenOfficeUtil {

    static Logger log = LoggerFactory.getLogger(OpenOfficeUtil.class);

    private static Process openOfficeProcess = null;

    private static OpenOfficeConnection connection = null;

    /**
     * 得到OpenOfficeConnection，使用后记得使用connection.disconnect()关闭连接
     * @author RunningHong at 2018/12/20 16:04
     * @return OpenOfficeConnection
     */
    public static OpenOfficeConnection getOpenOfficeConnection() {
        // OpenOffice的安装目录
        String OpenOffice_HOME = PropertiesUtil.getPropertyFromPropertyFile("report.properties", "OpenOffice_HOME");

        // 启动服务的命令
        String OpenOffice_command = PropertiesUtil.getPropertyFromPropertyFile("report.properties", "OpenOffice_command");

        // 启动OpenOffice的服务的完整命令
        String fullCommand = OpenOffice_HOME + OpenOffice_command;

        try {
            if (openOfficeProcess == null) {
                openOfficeProcess = Runtime.getRuntime().exec(fullCommand);
            }

            if (connection == null) { // connection为空创建连接
                connection = new SocketOpenOfficeConnection(8100);
            }

            connection.connect();
        } catch (Exception e) {
            log.error("OpenOffice连接获取失败，请检查OpenOffice服务是否启动！！！");
            e.printStackTrace();
        }
        return connection;
    }

    /**
     * 使用openOffice将excel流转换为PDF流并放于os中
     * @author RunningHong at 2018/12/19 20:42
     * @param
     * @return
     */
    public static void xlsStreamToPdfStream(InputStream is, OutputStream os) {
        // 得到OpenOffice的连接
        OpenOfficeConnection connection = OpenOfficeUtil.getOpenOfficeConnection();

        // 转换器
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);

        // xls格式
        DocumentFormat xlsFormat = new DocumentFormat("Microsoft Excel", DocumentFamily.SPREADSHEET, "application/vnd.ms-excel", "xls");
        xlsFormat.setExportFilter(DocumentFamily.SPREADSHEET, "MS Excel 97");

        // pdf格式
        DocumentFormat pdfFormat = new DocumentFormat("Portable Document Format","application/pdf", "pdf");
        pdfFormat.setExportFilter(DocumentFamily.DRAWING, "draw_pdf_Export");
        pdfFormat.setExportFilter(DocumentFamily.PRESENTATION, "impress_pdf_Export");
        pdfFormat.setExportFilter(DocumentFamily.SPREADSHEET, "calc_pdf_Export");
        pdfFormat.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");

        // 将xls流装换为pdf流
        converter.convert(is, xlsFormat, os, pdfFormat);
        connection.disconnect();
    }


    /**
     * 使用openOffice将excel流转换为HTML流并放于os中
     * @author RunningHong
     * @date 2018/12/7 17:56
     * @param
     */
    public static void excelStreamToHtmlStream(InputStream is, OutputStream os) {
        // 得到OpenOffice的连接
        OpenOfficeConnection connection = OpenOfficeUtil.getOpenOfficeConnection();

        // 转换器
        DocumentConverter converter = new OpenOfficeDocumentConverter(connection);

        // xls格式
        DocumentFormat xlsFormat = new DocumentFormat("Microsoft Excel", DocumentFamily.SPREADSHEET, "application/vnd.ms-excel", "xls");
        xlsFormat.setExportFilter(DocumentFamily.SPREADSHEET, "MS Excel 97");

        // html格式
        DocumentFormat htmlFormat = new DocumentFormat("HTML", DocumentFamily.TEXT, "text/html", "html");
        htmlFormat.setExportFilter(DocumentFamily.PRESENTATION, "impress_html_Export");
        htmlFormat.setExportFilter(DocumentFamily.SPREADSHEET, "HTML (StarCalc)");
        htmlFormat.setExportFilter(DocumentFamily.TEXT, "HTML (StarWriter)");

        final DocumentFormat pdf = new DocumentFormat("Portable Document Format", "application/pdf", "pdf");
        pdf.setExportFilter(DocumentFamily.DRAWING, "draw_pdf_Export");
        pdf.setExportFilter(DocumentFamily.PRESENTATION, "impress_pdf_Export");
        pdf.setExportFilter(DocumentFamily.SPREADSHEET, "calc_pdf_Export");
        pdf.setExportFilter(DocumentFamily.TEXT, "writer_pdf_Export");

        // 将xls流装换为html流
        converter.convert(is, xlsFormat, os, htmlFormat);
        connection.disconnect();
    }
}
