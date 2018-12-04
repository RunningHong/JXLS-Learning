package com.zh.jxls.util;

import com.zh.jxls.entity.Employee;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
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
 * 对Excel进行操作
 * @author RunningHong
 * @create 2018-12-03 20:59
 */
public class ExcelUtil {


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

            // 先生成Excel文件，excel内容信息存放于out流中
            this.generateExcelToStream(params, out);

            // 将Excel文件转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelToHtml(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
        } catch (Exception e) {
            throw new Exception("生成HTML文件出错，原因：" + e.getMessage(), e);
        }
    }

    /**
     * 生成Excel，生成的内容在OutputStream out中
     * @Author RunningHong
     * @Date 2018/12/3 19:50
     * @Param
     * @return
     */
    public void generateExcelToStream(Map<String, Object> params, OutputStream out) {

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
     * 读取Excel模板文件，生成报表Excel，存放到指定位置
     * @Author RunningHong
     * @Date 2018/12/4 15:36
     * @Param xlsTemplatePath：模板文件的位置， outPath:存放生成文件的位置
     * @return
     */
    public void generateExcelToFile(File xlsTemplatePath, File outPath) {
        // 生成临时数据
        List<Object> empList = generateDate();

        xlsTemplatePath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\src\\main\\resources\\templateFile\\测试报表.xls");

        SimpleDateFormat sdf = new SimpleDateFormat( "(MM_dd HH_mm)" );
        String timeFormat = sdf.format(new Date());
        String fileName = "temp" + timeFormat +".xls";

        outPath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\out\\generateTemp\\" + fileName);

        try {
            InputStream is = new FileInputStream(xlsTemplatePath);
            OutputStream os = new FileOutputStream(outPath);

            // 往JXLS的Context中添加数据
            Context context = new Context();
            context.putVar("emps", empList);

            // 使用JXLSHelper处理模板
            JxlsHelper.getInstance().processTemplate(is, os, context);
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

    /**
     * 判断是否时Excel2003版本
     * @Author RunningHong
     * @Date 2018/12/3 21:30
     * @Param
     * @return
     */
    public static boolean isExcel2003(InputStream is) {
        try {
            new HSSFWorkbook(is);
        } catch (Exception e) {
            return false;
        }
        return true;
    }


    /**
     * 判断是否时Excel2007版本
     * @Author RunningHong
     * @Date 2018/12/3 21:31
     * @Param
     * @return
     */
    public static boolean isExcel2007(InputStream is) {
        try {
            new XSSFWorkbook(is);
        } catch (Exception e) {
            return false;
        }
        return true;
    }
}
