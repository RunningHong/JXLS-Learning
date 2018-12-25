package com.zh.jxls.util;

import org.apache.commons.codec.binary.Base64;
import org.springframework.util.StringUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * 文件工具类
 * @author RunningHong
 * @since 2018-12-04 10:07
 */
public class FileUtil {

    /**
     * base64解码
     * @author RunningHong at 2018/12/12 10:05
     * @param picInfo 待解码信息
     * @return 图片解码后的byte[]
     */
    public static byte[] decodeBase64(String picInfo) {
        if (StringUtils.isEmpty(picInfo)) { // 解析空信息时
            return null;
        }
        // 传递过程中  "+" 变为了 " ",此处转换回来
        String base64Info = picInfo.replaceAll(" ", "+");

        // 数据中：data:image/png;base64,iVBORw0KGgoAAAANS... 在"base64,"之后的才是图片信息
        String[] baseArray = base64Info.split("base64,");

        // 使用了Apache commons codec的包来解析Base64
        byte[] picByteArr = Base64.decodeBase64(baseArray[1]);

        return picByteArr;
    }

    /**
     *保存byte数组流到指定文件
     * @author RunningHong at 2018/12/12 16:30
     * @param byteArr byte数组流
     * @param filePath 保存路径
     * @return
     */
    public static void saveByteArrayToFile(byte[] byteArr, File filePath) {
        OutputStream out = null;
        try {
            out = new FileOutputStream(filePath);
            out.write(byteArr);
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 获取文件的类型（扩展名）<br>
     * 如文件D:ddd.xls返回【.xls】
     * @author RunningHong
     * @date 2018/12/4 10:08
     * @param file 文件
     * @return 文件扩展名
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

    /**
     * 生成文件路径
     * @author RunningHong
     * @date 2018/12/6 10:54
     * @param filePath 需要生成的路径
     * @return
     */
    public static void makeFilePath(String filePath) {
        File filePathFile = new File(filePath);
        if (!filePathFile.exists()) {
            filePathFile.mkdirs();
        }
    }

}
