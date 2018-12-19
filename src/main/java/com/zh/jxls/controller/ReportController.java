package com.zh.jxls.controller;

import com.zh.jxls.util.ExportUtil;
import com.zh.jxls.util.FileUtil;
import com.zh.jxls.util.GenerateUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
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
     * @param response
     */
    @RequestMapping("/reportPreview")
    public void reportHtmlPreview(HttpServletResponse response) {
        Map<String, Object> params = new HashMap<>();
        GenerateUtil.generateHtmlToResponse(params, response);
    }

    /**
     * 获取并保存EChart图片
     * @author RunningHong at 2018/12/11 12:00
     * @param picInfo 图片信息base64加密过
     */
    @RequestMapping(value="/saveChart",method= RequestMethod.POST)
    public void saveChartImage(String picInfo) {
        // 图片解码后为byte[]
        byte[] picByteArr = FileUtil.decodeBase64(picInfo);

        // 图片保存路径
        String saveImagePath = "D:\\CodeSpace\\ideaWorkspace\\Report-Learning\\out\\generatePic\\chart.jpg";

        // 解码base64并把图片放到指定位置
        FileUtil.saveByteArrayToFile(picByteArr, new File(saveImagePath));
    }

    /**
     * 统计报表导出
     * @author RunningHong at 2018/12/19 21:18
     * @param
     */
    @RequestMapping("/reportExport")
    public void reportExport(HttpServletResponse response) {
        Map<String, Object> params = new HashMap<>();
        ExportUtil.exportMesToXls(params, response);
    }




}
