package com.zh.jxls;

import com.zh.jxls.entity.Employee;
import org.junit.Test;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.util.LinkedList;
import java.util.List;

public class JXLSTest {
    static Logger log = LoggerFactory.getLogger(JXLSTest.class);

    /**
     * 测试
     */
    @Test
    public void jxlsTest() {
        List<Object> empList = generateDate();

        File xlsTemplatePath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\src\\test\\resources\\templateFile\\测试报表.xls");
        File outPath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\out\\jxls_temp1.xls");

        try {
            InputStream is = new FileInputStream(xlsTemplatePath);
            OutputStream os = new FileOutputStream(outPath);


            Context context = new Context();
            context.putVar("emps", empList);


            JxlsHelper.getInstance().processTemplate(is, os, context);
        } catch (Exception e) {
            e.printStackTrace();
        }

    }

    /**
     * 数据生成
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
