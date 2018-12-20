package com.zh.jxls.util;

import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.Map;

/**
 * 用于前台预览展示
 * @author RunningHong
 * @create 2018-12-19 19:47
 */
public class PreviewUtil {

    static Logger log = LoggerFactory.getLogger(ExcelToHtmlUtil.class);

    /**
     * 生成HTML,文件流存于response中，用于前台展示
     * @Author RunningHong
     * @Date 2018/12/3 19:55
     * @Param
     * @return
     */
    public static void generateHtmlToResponse(Map<String, Object> params, HttpServletResponse response) {
        try {
            response.setContentType("text/html");
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // 参数处理,提供给jxls渲染模板
            Context context = ExcelUtil.handleParams(params);

            // Excel模板文件读取
            String templateFilePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "templateFilePath");
            InputStream tempXlsIs = new FileInputStream(new File(templateFilePath));

            // 使用jxsl渲染，渲染后的数据存放于out中
            JxlsHelper.getInstance().processTemplate(tempXlsIs, out, context);

            // 将Excel流存储到文件
            String saveExcelStreamToFilePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "saveExcelStreamToFilePath");
            ExcelUtil.excelStreamToFile(new ByteArrayInputStream(out.toByteArray()), new File(saveExcelStreamToFilePath));

            // 读取Excel流中的图片到指定位置
            String excelPictureSavePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "excelPictureSavePath");
            ExcelUtil.generateExcelPictureToFile(new ByteArrayInputStream(out.toByteArray()), excelPictureSavePath);


            // 使用OpenOffice将Excel文件流out转化为HTML，数据存储在response中
            // ExcelToHtmlUtil.excelStreamToHtmlStreamByOpenOffice(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());

            // 使用poi将Excel文件流out转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelStreamToHtmlStreamByPoi(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
        } catch (Exception e) {
            log.error("根据Excel模板文件生成预览HTML失败");
            e.printStackTrace();
        }
    }


}
