package com.zh.jxls.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
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
        String exportName = params.get("exportName").toString();
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
            response.setHeader("content-disposition","attachment;filename="+exportName+".xls");
            response.setCharacterEncoding("UTF-8");
            response.setContentType("application/vnd.ms-excel");

            // 生成信息
            PreviewUtil.previewHtml(params, response);
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
