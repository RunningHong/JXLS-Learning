package com.zh.jxls.controller;

import com.zh.jxls.util.ExportUtil;
import com.zh.jxls.util.FileUtil;
import com.zh.jxls.util.PreviewUtil;
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
        PreviewUtil.generateHtmlToResponse(params, response);
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

        if ("pdf".equals(exportType)) {
            // 导出为pdf格式
            ExportUtil.exportMesToPdfByOpenOffice(response);
        } else {
            // 导出为excel格式
            ExportUtil.exportMesToXls(params, response);
        }
    }





}
