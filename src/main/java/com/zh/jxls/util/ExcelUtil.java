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
     * 将2003Excel流转换为文件
     * @Author RunningHong
     * @Date 2018/12/8 13:52
     * @Param is:2003Excel流，outFile:输出文件
     * @return
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
     * 图片存放于Excel新建的一个TestPicSheet钟
     * @Author RunningHong
     * @Date 2018/12/6 14:44
     * @Param is:图片的流  workbook:带插入Excel的workbook
     * @return 
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
        log.info("图片流写入xls成功");
    }


    /**
     * 生成HTML文件,文件流存于response中
     * @Author RunningHong
     * @Date 2018/12/3 19:55
     * @Param
     * @return
     */
    public static void generateHtmlToResponse(Map<String, Object> params, HttpServletResponse response) {
        try {
            response.setContentType("text/html");
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // 参数处理,提供给jxls渲染模板
            Context context = handleParams(params);

            // Excel模板文件读取
            String templateFilePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "templateFilePath");
            InputStream tempXlsIs = new FileInputStream(new File(templateFilePath));

            // 使用jxsl渲染，渲染后的数据存放于out中
            JxlsHelper.getInstance().processTemplate(tempXlsIs, out, context);

            // 将Excel流存储到文件
            String saveExcelStreamToFilePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "saveExcelStreamToFilePath");
            ExcelUtil.excelStreamToFile(new ByteArrayInputStream(out.toByteArray()), new File(saveExcelStreamToFilePath));

            // 读取Excel流中的图片到指定位置
            String excelPictureSavePath = PropertiesUtil.getPropertyFromPropertyFile("filePath.properties", "excelPictureSavePath");
            ExcelUtil.generateExcelPictureToFile(new ByteArrayInputStream(out.toByteArray()), excelPictureSavePath);


            // 使用OpenOffice将Excel文件流out转化为HTML，数据存储在response中
            // ExcelToHtmlUtil.excelStreamToHtmlStreamByOpenOffice(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
            // 使用poi将Excel文件流out转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelStreamToHtmlStreamByPoi(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
        } catch (Exception e) {
            log.error("根据Excel模板文件生成预览HTML失败");
            e.printStackTrace();
        }
    }


    /**
     * 生成Excel2003文件的所有图片到指定位置
     * @Author RunningHong
     * @Date 2018/12/6 10:06
     * @Param is：Excel流信息  picSavePath：图片存放位置
     * @return
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
                // String saveName = "pic" + (i+1) + new SimpleDateFormat( "_MM_dd HH_mm" ).format(new Date()) + ".jpg";
                String saveName = "pic" + (i+1) + ".jpg";

                FileOutputStream fos = new FileOutputStream(new File(picSavePath + saveName));
                fos.write(picDate.getData());
                fos.close();
            }
        }
    }

    /**
     * 处理参数，供给jxls渲染模板
     * @Author RunningHong
     * @Date 2018/12/8 14:31
     * @Param
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
     * @Author RunningHong
     * @Date 2018/12/3 19:54
     * @Param
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
