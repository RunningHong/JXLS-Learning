package com.zh.jxls.controller;

import com.zh.jxls.util.FileUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import java.io.File;

/**
 * 图表控制器
 * @author RunningHong
 * @create 2018-12-11 11:57
 */
@RestController
@RequestMapping("/chart")
public class ChartController {
    Logger log = LoggerFactory.getLogger(ChartController.class);

    /**
     * 获取并保存EChart图片
     * @author RunningHong at 2018/12/11 12:00
     * @param picInfo 图片信息base64加密过
     * @return 
     */
    @RequestMapping(value="/saveChartImage",method=RequestMethod.POST)
    public void saveChartImage(String picInfo) {
        // 图片保存路径
        String saveImagePath = "D:\\CodeSpace\\ideaWorkspace\\Report-Learning\\out\\generatePic\\chart.jpg";

        // 图片解码后为byte[]
        byte[] picByteArr = FileUtil.decodeBase64(picInfo);

        // 解码base64并把图片放到指定位置
        FileUtil.saveByteArrayToFile(picByteArr, new File(saveImagePath));

        log.info("保存Echarts图片文件成功。");
    }


}
