package com.zh.jxls.util;

import com.zh.jxls.entity.Employee;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * @author RunningHong
 * @create 2018-12-03 19:45
 */
public class JxlsUtil {

    /**
     * 生成HTML文件
     * @Author RunningHong
     * @Date 2018/12/3 19:55
     * @Param
     * @return
     */
    public void generateHtml(Map<String, Object> params, HttpServletResponse response) throws Exception {
        try {
            response.setContentType("text/html");
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // 先生成Excel文件
            this.generateExcel(params, out);

            // 将Excel文件转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelToHtml(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
        } catch (Exception e) {
            throw new Exception("生成HTML文件出错，原因：" + e.getMessage(), e);
        }
    }
    
    /**
     * 生成Excel，生成的内容在out中
     * @Author RunningHong
     * @Date 2018/12/3 19:50 
     * @Param 
     * @return 
     */
    public void generateExcel(Map<String, Object> params, OutputStream out) {

        // Excel模板文件读取位置
        File xlsTemplatePath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\src\\main\\resources\\templateFile\\测试报表.xls");

        SimpleDateFormat sdf = new SimpleDateFormat( "(MM_dd HH_mm)" );
        String timeFormat = sdf.format(new Date());
        String fileName = "temp" + timeFormat +".xls";

        // 生成测试数据
        List<Object> empList = generateDate();

        // 往context中添加数据
        Context context = new Context(params);
        context.putVar("emps", empList);

        try {
            InputStream is = new FileInputStream(xlsTemplatePath);
            JxlsHelper.getInstance().processTemplate(is, out, context);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 生成测试数据
     * @Author RunningHong
     * @Date 2018/12/3 19:54 
     * @Param 
     * @return 
     */
    public List<Object> generateDate() {
        List<Object> list = new LinkedList<>();

        for (int i = 0; i < 10; i++) {
            Employee emp = new Employee();
            emp.setName("测试人员" + i);
            emp.setAge(i+20);
            emp.setCode(100 + i + "");
            emp.setLinkTel("" + i + "8888");

            list.add(emp);
        }

        return list;
    }
}
