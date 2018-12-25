package com.zh.jxls.controller;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.zh.jxls.util.ExportUtil;
import com.zh.jxls.util.PreviewUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.Map;

/**
 * 报表控制器
 * @author RunningHong
 * @since 2018-12-03 20:12
 */
@RestController
@RequestMapping( value = "/reportController")
public class ReportController {

    /**
     * 统计报表HTML预览
     * @author RunningHong at 2018/12/3 20:13
     */
    @RequestMapping("/reportPreview")
    public void reportHtmlPreview(HttpServletResponse response) {
        Map<String, Object> params = new HashMap<>();
        PreviewUtil.previewHtml(params, response);
    }

    /**
     * 导出所有charts到Excel
     * 前台需要传需要导出的图片信息
     * @author RunningHong at 2018/12/25 20:15
     */
    @RequestMapping(value="/exportAllCharts",method= RequestMethod.POST)
    public void exportAllCharts(HttpServletRequest request, HttpServletResponse response) {
        JSONObject jsonObject = JSON.parseObject(request.getParameter("chartsJson"));
        JSONArray chartsArray = jsonObject.getJSONArray("chartsJsonArray");

        Map<String, Object> params = new HashMap<>();
        // 设置导出名称
        params.put("exportName", "统计图表");

        ExportUtil.exportChartToXls(params, chartsArray, response);
    }

    /**
     * 统计报表导出 <br>
     * 说明：<br>
     * 1)导出pdf需要使用OpenOffice服务，需要手动打开<br>
     * 2)导出pdf格式可能会出现列数过多导致折页现象，此时在excel模板文件中，把打印设置设为缩放打印（将所有列调整为一页）<br>
     * @author RunningHong at 2018/12/19 21:18
     * @param exportType 导出类型，exportType为pdf导出pdf格式，其余导出excel格式
     */
    @RequestMapping("/reportExport")
    public void reportExport(HttpServletResponse response, String exportType) {
        Map<String, Object> params = new HashMap<>();

        // 设置报表导出名称
        params.put("exportName", "统计报表");

        if ("pdf".equals(exportType)) {
            // 导出为pdf格式
            ExportUtil.exportMesToPdf(params, response);
        } else {
            // 导出为excel格式
            ExportUtil.exportMesToXls(params, response);
        }
    }





}
