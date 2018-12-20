package com.zh.jxls.util;

import com.artofsolving.jodconverter.openoffice.connection.OpenOfficeConnection;
import com.artofsolving.jodconverter.openoffice.connection.SocketOpenOfficeConnection;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 需要手动配置OpenOffice的位置
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

    private static OpenOfficeConnection connection = null;

    /**
     * 得到OpenOfficeConnection，使用后记得使用connection.disconnect()关闭连接
     * @author RunningHong at 2018/12/20 16:04
     * @return OpenOfficeConnection
     */
    public static OpenOfficeConnection getOpenOfficeConnection() {
        // 启动OpenOffice的服务的完整命令
        String fullCommand = OpenOffice_HOME + command;

        try {
            if (openOfficeProcess == null) {
                openOfficeProcess = Runtime.getRuntime().exec(fullCommand);
            }

            if (connection == null) { // connection为空创建连接
                // 创建OpenOffic连接
                connection = new SocketOpenOfficeConnection(8100);
            }

            connection.connect();
        } catch (Exception e) {
            log.error("OpenOffice连接获取失败，请检查OpenOffice服务是否启动！！！");
            e.printStackTrace();
        }

        return connection;
    }
}
