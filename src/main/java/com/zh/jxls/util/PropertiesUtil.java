package com.zh.jxls.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.Properties;

/**
 * Properties工具类
 * @author RunningHong
 * @create 2018-12-08 16:21
 */
public class PropertiesUtil {
    static Logger log = LoggerFactory.getLogger(PropertiesUtil.class);

    /**
     * 从特定的配置文件读取配置参数的值
     * 方法使用UTF-8格式读取，可以读取中文参数值
     * @Author RunningHong
     * @Date 2018/12/10 10:35
     * @Param propertyFilePath：配置文件相对路径    proName:参数名称
     * @return 参数值
     */
    public static String getPropertyFromPropertyFile(String propertyFilePath, String proName) {
        Properties properties = new Properties();
        try {
            InputStream is = ClassLoader.getSystemResourceAsStream(propertyFilePath);
            BufferedReader bf = new BufferedReader( new InputStreamReader(is));
            properties.load(bf);
            return properties.getProperty(proName);
        } catch (IOException e) {
            log.error("根据Properties文件 " + propertyFilePath + "  获取参数 " + proName + " 失败！");
            e.printStackTrace();
        }
        return properties.getProperty(proName);
    }

}
