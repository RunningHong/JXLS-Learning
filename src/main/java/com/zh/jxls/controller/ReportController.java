package com.zh.jxls.controller;

import com.zh.jxls.util.ExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.util.HashMap;
import java.util.Map;

/**
 * 报表控制器
 * @author RunningHong
 * @create 2018-12-03 20:12
 */
@RestController
@RequestMapping( value = "/reportController")
public class ReportController {

    /**
     * 统计报表HTML预览
     * @Author RunningHong
     * @Date 2018/12/3 20:13
     * @Param
     * @return
     */
    @RequestMapping("/reportHtmlPreview")
    public void reportHtmlPreview(HttpServletResponse response) {
        Map<String, Object> params = new HashMap<>();
        ExcelUtil.generateHtmlToResponse(params, response);
    }
}
