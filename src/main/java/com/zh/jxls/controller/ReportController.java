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
     * 统计报表导出
     * 说明：<br>
     * 1)导出pdf需要使用OpenOffice服务，需要手动打开<br>
     * 2)导出pdf格式可能会出现列数过多导致折页现象，此时在excel模板文件中，把打印设置设为缩放打印（将所有列调整为一页）<br>
     * @author RunningHong at 2018/12/26 15:02
     */
    @RequestMapping(value="/reportExport", method= RequestMethod.POST)
    public void reportExport(HttpServletRequest request, HttpServletResponse response) {
        ExportUtil.reportExport(request, response);
    }



}
