package com.zh.jxls.util;

import com.alibaba.fastjson.JSONArray;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;
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
     * 导出名称处理
     * @author RunningHong at 2018/12/21 11:26
     * @param params 里面包含导出名称exportName
     * @return 报表名称
     */
    public static String handleExportName(Map<String, Object> params) {
        String exportName = "";
        if(!params.containsKey("exportName")) { // 不存在key
            exportName = "默认名称";
        } else {
            exportName = params.get("exportName").toString();
        }
        try {
            // 导出文件名称
            if ("".equals(exportName) || (exportName == null)) {
                exportName = "默认名称";
            }

            exportName = URLEncoder.encode(exportName, "UTF-8");
        } catch (Exception e) {
            e.printStackTrace();
        }
        return exportName;
    }

    /**
     * 以xls格式导出统计信息
     * @author RunningHong at 2018/12/14 20:41
     * @param
     */
    public static void exportMesToXls(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 导出文件名称
            String exportName = ExportUtil.handleExportName(params);

            // 设置导出文件格式
            response.setHeader("content-disposition","attachment;filename=" + exportName + ".xls");
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");

            // 使用jxls渲染数据
            ByteArrayOutputStream out = JxlsUtil.excelLoadingToStreamByJxls(params);

            // 使用poi将Excel文件流out转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelStreamToHtmlStreamByPoi(new ByteArrayInputStream(out.toByteArray()),
                                                         response.getOutputStream());

        } catch (Exception e) {
            log.error("以XLS格式导出统计信息失败！");
            e.printStackTrace();
        }
    }

    /**
     * 以xls格式导出统计Chart
     * @author RunningHong at 2018/12/14 20:41
     * @param params 参数信息
     * @param chartsArray 图片的jsonArray信息
     * @param response
     * @return
     */
    public static void exportChartToXls(Map<String, Object> params, JSONArray chartsArray, HttpServletResponse response) {
        // 导出文件名称
        String exportName = ExportUtil.handleExportName(params);

        // 设置导出文件格式
        response.setHeader("content-disposition","attachment;filename="+exportName+".xls");
        response.setCharacterEncoding("UTF-8");
        response.setContentType("application/vnd.ms-excel");

        HSSFWorkbook workbook = new HSSFWorkbook();

        // 处理图片写入excel
        ExcelUtil.writePictureToExcel(chartsArray, workbook);

        try { // 将excel信息写入response
            OutputStream responseOs = response.getOutputStream();
            workbook.write(responseOs);
            responseOs.close();
        } catch (Exception e) {
            log.error("导出图表到Excel失败！");
            e.printStackTrace();
        }

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
    public static void exportMesToPdf(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 导出文件名称
            String exportName = ExportUtil.handleExportName(params);

            //设置页面编码格式
            response.setContentType("text/plain;charaset=utf-8");
            response.setContentType("application/vnd.ms-pdf");
            response.setHeader("Content-Disposition", "attachment; filename=" + exportName + ".pdf");

            // 使用jxls渲染数据
            ByteArrayOutputStream out = JxlsUtil.excelLoadingToStreamByJxls(params);

            // 使用OpenOffice把excel转换为PDF
            OpenOfficeUtil.xlsStreamToPdfStream(new ByteArrayInputStream(out.toByteArray()),
                                                response.getOutputStream());
        } catch (Exception e) {
            log.error("以PDF格式导出统计信息失败！");
            e.printStackTrace();
        }

    }




}
