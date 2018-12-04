package com.zh.jxls.util;

import java.io.File;

/**
 * 文件工具类
 * @author RunningHong
 * @create 2018-12-04 10:07
 */
public class FileUtil {

    /**
     * 获取文件的类型（扩展名）<br>
     * 如文件D:ddd.xls返回【.xls】
     * @Author RunningHong
     * @Date 2018/12/4 10:08
     * @Param
     * @return
     */
    public static String getFileType(File file) throws Exception{
        if(!file.exists()) {
            throw new Exception("文件不存在。");
        }

        String fileName = file.getName();

        // 获取文件的类型
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        return fileType;
    }
}
