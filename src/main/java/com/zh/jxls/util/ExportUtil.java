package com.zh.jxls.util;

import com.artofsolving.jodconverter.DocumentConverter;
import com.artofsolving.jodconverter.DocumentFamily;
import com.artofsolving.jodconverter.DocumentFormat;
import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.converter.OpenOfficeDocumentConverter;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
import java.util.HashMap;
import java.util.Map;

/**
 * 报表导出相关工具类
 * 导出为XLS，PDF
 * @author RunningHong
 * @since 2018-12-14 17:20
 */
public class ExportUtil {

    static Logger log = LoggerFactory.getLogger(ExportUtil.class);

    /**
     * 以xls格式导出统计信息
     * @author RunningHong at 2018/12/14 20:41
     * @param
     * @return
     */
    public static void exportMesToXls(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 导出文件名称
            String exportName = "报表导出文件";
            exportName = URLEncoder.encode(exportName, "UTF-8");

            // 设置导出文件格式
            response.setHeader("content-disposition","attachment;filename="+exportName+".xls");
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");

            // 生成信息
            PreviewUtil.generateHtmlToResponse(params, response);
        } catch (Exception e) {
            log.error("以XLS格式导出统计信息失败！");
            e.printStackTrace();
        }
    }

    /**
     * 以xls格式导出统计Chart
     * @author RunningHong at 2018/12/14 20:41
     * @param
     * @return
     */
    public static void exportChartToXls() {
        // TODO: 以xls格式导出统计Chart
    }

    /**
     * 以xls格式合并导出统计信息和统计Chart
     * @author RunningHong at 2018/12/19 20:03
     * @param
     * @return
     */
    public static void exportMesAndChartToXls() {
        // TODO: 以xls格式导出统计信息和统计Chart
    }

    /**
     * 以PDF格式导出统计信息,使用需要开启OpenOffice服务
     * @author RunningHong at 2018/12/20 17:05
     * @param
     * @return
     */
    public static void exportMesToPdfByOpenOffice(HttpServletResponse response) {
        try {
            // 导出文件名称
            String exportName = "报表导出文件";
            exportName = URLEncoder.encode(exportName, "UTF-8");

            //设置页面编码格式
            response.setContentType("text/plain;charaset=utf-8");
            response.setContentType("application/vnd.ms-pdf");
            response.setHeader("Content-Disposition", "attachment; filename=" + exportName + ".pdf");




            // 对信息进行处理
            Map<String, Object> params = new HashMap<>();
            // 参数处理,提供给jxls渲染模板
            Context context = ExcelUtil.handleParams(params);

            // Excel模板文件读取
            String templateFilePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "templateFilePath");
            InputStream tempXlsIs = new FileInputStream(new File(templateFilePath));

            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // 使用jxls渲染，渲染后的数据存放于out中
            JxlsHelper.getInstance().processTemplate(tempXlsIs, out, context);




            // 转换为PDF
            ExportUtil.xlsStreamToPdfStream(new ByteArrayInputStream(out.toByteArray()),
                                            response.getOutputStream());
        } catch (Exception e) {
            log.error("以PDF格式导出统计信息失败！");
            e.printStackTrace();
        }

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


}
