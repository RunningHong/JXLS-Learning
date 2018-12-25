package com.zh.jxls.util;

import com.alibaba.fastjson.JSONArray;
import com.zh.jxls.entity.Employee;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.*;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.imageio.ImageIO;
import javax.servlet.http.HttpServletResponse;
import java.awt.image.BufferedImage;
import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 对Excel进行操作
 * @author RunningHong
 * @since 2018-12-03 20:59
 */
public class ExcelUtil {

    static Logger log = LoggerFactory.getLogger(ExcelUtil.class);


    /**
     * 将03版Excel流保存到Excel文件
     * @author RunningHong
     * @date 2018/12/8 13:52
     * @param is 2003Excel流
     * @param outFile 输出文件
     */
    public static void excelStreamToFile(InputStream is, File outFile) {
        try {
            HSSFWorkbook workBook = new HSSFWorkbook(is);
            workBook.write(outFile);
            log.info("Excel流转换xls成功，生成文件路径:" + outFile.getAbsolutePath());
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * 将图片的流is写入Excel的workbook
     * 图片存放于Excel新建的一个sheet中
     * @author RunningHong
     * @date 2018/12/6 14:44
     * @param chartsArray 图片的JSONArray
     * @param workbook 待插入Excel的workbook
     */
    public static void writePictureToExcel(JSONArray chartsArray, HSSFWorkbook workbook){
        try {
            // 循环插入图片
            for (int i = 0; i < chartsArray.size(); i++) {
                String sheetName = "统计图表" + (i+1);
                HSSFSheet createSheet = workbook.createSheet(sheetName);

                //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
                HSSFPatriarch patriarch = createSheet.createDrawingPatriarch();

                // 设置图片坐标
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255,(short) 1, 1, (short) 5, 8);

                // 设置图片DONT_MOVE_AND_RESIZE
                anchor.setAnchorType(ClientAnchor.AnchorType.byId(3));

                // 图片解码为byte[]
                byte[] picByteArray = FileUtil.decodeBase64((String)chartsArray.get(i));

                // 往workbook中插入图片
                patriarch.createPicture(anchor, workbook.addPicture(picByteArray, HSSFWorkbook.PICTURE_TYPE_PNG));
            }
        } catch (Exception e) {
            log.error("图片写入excel失败！");
            e.printStackTrace();
        }
    }



    /**
     * 取出03版Excel文件的所有图片，并保存到指定位置
     * @author RunningHong
     * @date 2018/12/6 10:06
     * @param is Excel流信息
     * @param picSavePath 图片存放位置
     */
    public static void generateExcelPictureToFile(InputStream is, String picSavePath) throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook(is);

        File picSaveFile = new File(picSavePath);
        if (!picSaveFile.exists()) {
            picSaveFile.mkdirs();
        }

        List<HSSFPictureData> pics = workbook.getAllPictures();
        if (pics != null) {
            for (int i = 0; i < pics.size(); i++) {
                HSSFPictureData  picDate = (HSSFPictureData) pics.get(i);

                String saveName = "pic" + (i+1) + ".jpg";
                FileOutputStream fos = new FileOutputStream(new File(picSavePath + saveName));
                fos.write(picDate.getData());
                fos.close();
            }
        }
    }


}
