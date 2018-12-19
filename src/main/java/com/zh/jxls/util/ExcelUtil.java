package com.zh.jxls.util;

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
 * @create 2018-12-03 20:59
 */
public class ExcelUtil {

    static Logger log = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 将xls转换为PDF
     * @author RunningHong at 2018/12/19 22:09
     * @param 
     * @return 
     */
    public static void xlsToPdf() {

    }

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
     * @param is 图片的流
     * @param workbook 带插入Excel的workbook
     */
    public static void writePictureToExcel(InputStream is, HSSFWorkbook workbook) throws Exception {
        // 创建新的sheet，用于存放图片sheet
        HSSFSheet createSheet = workbook.createSheet("TestPicSheet");

        //画图的顶级管理器，一个sheet只能获取一个（一定要注意这点）
        HSSFPatriarch patriarch = createSheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 255, 255,(short) 1, 1, (short) 5, 8);

        // 设置图片DONT_MOVE_AND_RESIZE
        anchor.setAnchorType(ClientAnchor.AnchorType.byId(3));

        // 读取图片io
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg = ImageIO.read(is);
        ImageIO.write(bufferImg, "jpg", byteArrayOut);

        // 往workbook中插入图片
        patriarch.createPicture(anchor, workbook.addPicture(byteArrayOut.toByteArray(),
                                                            HSSFWorkbook.PICTURE_TYPE_JPEG));
        log.info("图片流写入xls成功。");
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

    /**
     * 处理参数，供给jxls渲染模板
     * @author RunningHong
     * @date 2018/12/8 14:31
     * @param params 待处理的参数
     * @return
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
