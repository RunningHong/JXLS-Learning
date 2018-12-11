package com.zh.jxls.controller;

import org.apache.commons.codec.binary.Base64;
import org.apache.commons.compress.utils.IOUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

/**
 * 图表控制器
 * @author RunningHong
 * @create 2018-12-11 11:57
 */
@RestController
@RequestMapping("/chart")
public class ChartController {

    /**
     * 获取并保存EChart图片到本地
     * @author RunningHong at 2018/12/11 12:00
     * @param picInfo 图片信息
     * @return 
     */
    @RequestMapping(value="/saveImage",method= RequestMethod.POST)
    public void saveImage(String picInfo) {
        System.out.println("开始保存Echarts图片文件。");

        String saveImagePath = "D:\\CodeSpace\\ideaWorkspace\\Report-Learning\\out\\generatePic\\chart.jpg";

        // 传递过程中  "+" 变为了 " ".
        String newPicInfo = picInfo.replaceAll(" ", "+");

        // 解码base64并把图片放到指定位置
        decodeBase64(newPicInfo, new File(saveImagePath));
        System.out.println("保存Echarts图片文件成功。");
    }

    /**
     * 解析Base64位信息并输出到某个目录下面.
     * @param base64Info base64串
     * @param picPath 生成的文件路径
     * @return 文件地址
     */
    private File decodeBase64(String base64Info, File picPath) {
        if (StringUtils.isEmpty(base64Info)) {
            return null;
        }

        // 数据中：data:image/png;base64,iVBORw0KGgoAAAANS... 在"base64,"之后的才是图片信息
        String[] arr = base64Info.split("base64,");

        // 将图片输出到系统某目录.
        OutputStream out = null;
        try {
            // 使用了Apache commons codec的包来解析Base64
            byte[] buffer = Base64.decodeBase64(arr[1]);
            out = new FileOutputStream(picPath);
            out.write(buffer);
        } catch (IOException e) {
            e.printStackTrace();
            //log.error("解析Base64图片信息并保存到某目录下出错!", e);
        } finally {
            IOUtils.closeQuietly(out);
        }

        return picPath;
    }
}
