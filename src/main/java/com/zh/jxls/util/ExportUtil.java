package com.zh.jxls.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.UnsupportedEncodingException;
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
            GenerateUtil.generateHtmlToResponse(params, response);
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
     * 以PDF格式导出统计信息
     * @author RunningHong at 2018/12/19 20:42
     * @param
     * @return
     */
    public static void exportMesToPdf(Map<String, Object> params, HttpServletResponse response) {
        try {
            // 导出文件名称
            String exportName = "报表导出文件";
            exportName = URLEncoder.encode(exportName, "UTF-8");

            //设置页面编码格式
            response.setContentType("text/plain;charaset=utf-8");
            response.setContentType("application/vnd.ms-pdf");
            response.setHeader("Content-Disposition", "attachment; filename=" + exportName + ".pdf");


        } catch (Exception e) {
            log.error("以PDF格式导出统计信息失败！");
            e.printStackTrace();
        }


        // TODO: 以PDF格式导出统计信息
    }


}
