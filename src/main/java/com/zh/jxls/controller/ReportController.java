package com.zh.jxls.controller;

import com.zh.jxls.util.ExcelUtil;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.text.SimpleDateFormat;
import java.util.Date;
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
    public void reportHtmlPreview(HttpServletResponse response) throws Exception {
        testGenerateExcelToFile();

        ExcelUtil excelUtil = new ExcelUtil();
        Map<String, Object> params = new HashMap<>();
        excelUtil.generateHtmlToResponse(params, response);

    }

    /**
     * 测试jxls根据模板文件生成Excel文件
     * @Author RunningHong
     * @Date 2018/12/5 11:39
     * @Param
     * @return
     */
    public void testGenerateExcelToFile() {
        // 模板文件路径
        String xlsTempPath = "D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\src\\main\\resources\\templateFile\\";
        String xlsTempName = "测试报表.xls";
        File xlsTemplateFile = new File(xlsTempPath + xlsTempName);

        // 导出文件路径
        String outFilePath = "D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\out\\generateTemp\\";
        String outFileName = "temp"  + new SimpleDateFormat( "_MM_dd HH_mm" ).format(new Date()) + ".xls";
        File outFile = new File(outFilePath + outFileName);

        ExcelUtil.generateExcelToFile(xlsTemplateFile, outFile);
    }
}
