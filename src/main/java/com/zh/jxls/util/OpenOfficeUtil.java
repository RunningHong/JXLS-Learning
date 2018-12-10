package com.zh.jxls.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author RunningHong
 * @create 2018-12-07 17:00
 */
public class OpenOfficeUtil {

    static Logger log = LoggerFactory.getLogger(OpenOfficeUtil.class);

    // OpenOffice的安装目录
    private static String OpenOffice_HOME = "C:\\Program Files (x86)\\OpenOffice 4\\program\\";
    // "C:\Program Files (x86)\OpenOffice 4\program\soffice.exe"

    // 启动服务的命令
    private static String command = "soffice -headless -accept=\"socket,host=127.0.0.1,port=8100;urp;\"";

    private static Process openOfficeProcess = null;

    /**
     * 启动openOffice服务
     */
    public static Process startOpenOffice() {
        // 启动OpenOffice的服务的完整命令
        String fullCommand = OpenOffice_HOME + command;

        try {
            if (openOfficeProcess == null) {
                openOfficeProcess = Runtime.getRuntime().exec(fullCommand);
            }
            log.info("OpenOffice开启成功");
        } catch (Exception e) {
            e.printStackTrace();
        }

        return openOfficeProcess;
    }
}
