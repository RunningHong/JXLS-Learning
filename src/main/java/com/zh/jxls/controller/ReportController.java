package com.zh.jxls.controller;

import com.zh.jxls.util.JxlsUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;

/**
 * 报表控制器
 * @author RunningHong
 * @create 2018-12-03 20:12
 */
@Controller
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
    public void reportHtmlPreview(HttpServletResponse response) throws Exception {
        JxlsUtil jxlsUtil = new JxlsUtil();
        jxlsUtil.generateHtml(null, response);
    }
}
