package com.zh.jxls.util;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.OutputStream;
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
     * 处理前台参数
     * @author RunningHong at 2018/12/26 9:50
     * @param jsonParams param的json信息
     * @return Map<String, Object>
     */
    public static Map<String, Object> handleParams(JSONObject jsonParams) {
        return jsonParams;
    }

    /**
     * 导出名称处理
     * @author RunningHong at 2018/12/21 11:26
     * @param params 里面包含导出名称exportName
     * @return 报表名称
     */
    public static String handleExportName(Map<String, Object> params) {
        String exportName;
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
     * 统计报表导出
     * @author RunningHong at 2018/12/26 16:42
     */
    public static void reportExport(HttpServletRequest request, HttpServletResponse response) {
        // 获取前台参数
        JSONObject jsonParams = JSON.parseObject(request.getParameter("params"));
        Map<String, Object> paramsMap = ExportUtil.handleParams(jsonParams);

        // 设置报表导出名称
        paramsMap.put("exportName", "导出测试");

        // 导出图表信息
        JSONArray chartsMes = (JSONArray) paramsMap.get("chartsArray");

        // 导出类型【信息导出、图表导出、合并导出】
        String exportType = (String)paramsMap.get("exportType");

        // 导出文件类型【xls, pdf】
        String exportFileType = (String)paramsMap.get("exportFileType");

        if ("信息导出".equals(exportType) && "xls".equals(exportFileType)) { // 仅导出信息到xls
            ExportUtil.exportMesToXls(paramsMap, response);

        } else if ("图表导出".equals(exportType) && "xls".equals(exportFileType)) { // 仅导出charts到xls
            ExportUtil.exportChartToXls(paramsMap, chartsMes, response);

        } else if ("合并导出".equals(exportType) && "xls".equals(exportFileType)) { // 合并导出，同时导出信息和图表信息到xls
            ExportUtil.exportMesAndChartToXls(paramsMap, chartsMes, response);

        } else if ("信息导出".equals(exportType) && "pdf".equals(exportFileType)) { // 仅导出信息到pdf
            ExportUtil.exportMesToPdf(paramsMap, response);

        } else {
            log.error("系统暂不支持此类型导出！");
        }
    }

    /**
     * 设置文件导出格式
     * @author RunningHong at 2018/12/26 15:15
     * @param params 导出所需参数
     * @param exportFileType 导出文件类型【xls,pdf,...】
     * @return
     */
    public static HttpServletResponse setExportType(Map<String, Object> params, String exportFileType, HttpServletResponse response) {
        // 获取导出文件名称
        String exportName = ExportUtil.handleExportName(params);

        if ("xls".equals(exportFileType)) {
            // 设置导出文件格式
            response.setHeader("content-disposition","attachment;filename=" + exportName + ".xls");
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");
        } else if ("pdf".equals(exportFileType)) {
            //设置页面编码格式
            response.setContentType("text/plain;charaset=utf-8");
            response.setContentType("application/vnd.ms-pdf");
            response.setHeader("Content-Disposition", "attachment; filename=" + exportName + ".pdf");
        } else {
            log.error("报表暂不支持此类型导出！");
        }
        return response;
    }
    
    

    /**
     * 以xls格式导出统计信息
     * @author RunningHong at 2018/12/14 20:41
     */
    public static void exportMesToXls(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 设置文件导出格式
            response = setExportType(params, "xls", response);

            // 使用jxls渲染数据
            ByteArrayOutputStream baos = JxlsUtil.excelLoadingToStreamByJxls(params);

            // 将信息存入response
            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(baos.toByteArray());
            HSSFWorkbook workbook = new HSSFWorkbook(byteArrayInputStream);
            OutputStream os = response.getOutputStream();
            workbook.write(os);
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
     */
    public static void exportChartToXls(Map<String, Object> params, JSONArray chartsArray, HttpServletResponse response) {
        try {
            // 设置文件导出格式
            response = setExportType(params, "xls", response);

            HSSFWorkbook workbook = new HSSFWorkbook();
            // 处理图片写入excel
            ExcelUtil.writePictureToExcel(chartsArray, workbook);

            // 图片信息存入response
            OutputStream os = response.getOutputStream();
            workbook.write(os);
        } catch (Exception e) {
            log.error("导出图表到Excel失败！");
            e.printStackTrace();
        }

    }

    /**
     * 以xls格式合并导出统计信息和统计Chart
     * @author RunningHong at 2018/12/19 20:03
     * @param
     */
    public static void exportMesAndChartToXls(Map<String, Object> params, JSONArray chartsArray, HttpServletResponse response) {
        try {
            // 设置文件导出格式
            response = setExportType(params, "xls", response);

            // 使用jxls渲染数据
            ByteArrayOutputStream byos = JxlsUtil.excelLoadingToStreamByJxls(params);

            // 将信息存入response
            ByteArrayInputStream byteArrayInputStream = new ByteArrayInputStream(byos.toByteArray());
            HSSFWorkbook workbook = new HSSFWorkbook(byteArrayInputStream);

            // 图片写入excel
            ExcelUtil.writePictureToExcel(chartsArray, workbook);

            // 图片信息写入response
            OutputStream os = response.getOutputStream();
            workbook.write(os);
        } catch (Exception e) {
            log.error("合并导出失败！");
            e.printStackTrace();
        }
    }

    /**
     * 以PDF格式导出统计信息,使用需要开启OpenOffice服务
     * @author RunningHong at 2018/12/20 17:05
     * @param params 相关参数
     */
    public static void exportMesToPdf(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 设置文件导出格式
            response = setExportType(params, "pdf", response);

            // 使用jxls渲染数据
            ByteArrayOutputStream byos = JxlsUtil.excelLoadingToStreamByJxls(params);

            // 使用OpenOffice把excel转换为PDF
            OpenOfficeUtil.xlsStreamToPdfStream(new ByteArrayInputStream(byos.toByteArray()),
                                                response.getOutputStream());
        } catch (Exception e) {
            log.error("以PDF格式导出统计信息失败！");
            e.printStackTrace();
        }
    }




}
