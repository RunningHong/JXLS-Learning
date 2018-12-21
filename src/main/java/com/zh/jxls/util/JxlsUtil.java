package com.zh.jxls.util;

import com.zh.jxls.entity.Employee;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * JXLS工具类，根据模板渲染excel文件
 * @author RunningHong
 * @since 2018-12-21 10:47
 */
public class JxlsUtil {
    static Logger log = LoggerFactory.getLogger(JxlsUtil.class);

    /**
     * 使用jxls把文件加载到流中
     * @author RunningHong at 2018/12/21 10:53
     * @param
     * @return
     */
    public static ByteArrayOutputStream excelLoadingToStreamByJxls(Map<String, Object> params) {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();

        try {
            // 参数处理,提供给jxls渲染模板
            Context context = JxlsUtil.handleParams(params);

            // Excel模板文件读取
            String templateFilePath = PropertiesUtil.getPropertyFromPropertyFile("report.properties", "templateFilePath");
            InputStream tempXlsIs = new FileInputStream(new File(templateFilePath));

            // 使用jxsl渲染，渲染后的数据存放于out中
            JxlsHelper.getInstance().processTemplate(tempXlsIs, byteArrayOutputStream, context);
        } catch (IOException e) {
            log.error("JXLS渲染模板文件失败！！！");
            e.printStackTrace();
        }

        return byteArrayOutputStream;
    }

    /**
     * 处理参数，供给jxls渲染模板
     * @author RunningHong
     * @date 2018/12/8 14:31
     * @param params 待处理的参数
     * @return Context
     */
    public static Context handleParams(Map<String, Object> params) {
        // 生成测试数据
        List<Object> empList = generateDate();

        // 往context中添加数据
        Context context = new Context(params);
        context.putVar("emps", empList);

        return context;
    }

    /**
     * 生成测试数据
     * @author RunningHong
     * @date 2018/12/3 19:54
     * @param
     * @return
     */
    public static List<Object> generateDate() {
        List<Object> list = new LinkedList<>();

        String time = new SimpleDateFormat( "(YY_MM_dd HH_mm_ss)" ).format(new Date());
        for (int i = 0; i < 10; i++) {
            Employee emp = new Employee();
            emp.setName("测试人员" + i + "\r\n" + "iiiiiiiiiiiiiiiiii");
            emp.setAge(i+20);
            emp.setCode(100 + i + "");
            emp.setLinkTel(time);

            list.add(emp);
        }
        return list;
    }





}
