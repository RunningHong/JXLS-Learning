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
import java.util.Date;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

/**
 * 对Excel进行操作
 * @author RunningHong
 * @create 2018-12-03 20:59
 */
public class ExcelUtil {

    static Logger log = LoggerFactory.getLogger(ExcelUtil.class);

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
    }


    /**
     * 生成HTML文件,文件流存于response中
     * @Author RunningHong
     * @Date 2018/12/3 19:55
     * @Param
     * @return
     */
    public static void generateHtmlToResponse(Map<String, Object> params, HttpServletResponse response) throws Exception {
        try {
            response.setContentType("text/html");
            ByteArrayOutputStream out = new ByteArrayOutputStream();

            // 读取模板文件，使用jxsl渲染，渲染后的数据存放于out中
            generateExcelToStream(params, out);

            // 将Excel文件流out转化为HTML，数据存储在response中
            ExcelToHtmlUtil.excelToHtml(new ByteArrayInputStream(out.toByteArray()), response.getOutputStream());
        } catch (Exception e) {
            throw new Exception("生成HTML文件出错，原因：" + e.getMessage(), e);
        }
    }


    /**
     * 读取模板文件，使用jxsl渲染，渲染后的数据存放于out中
     * @Author RunningHong
     * @Date 2018/12/3 19:50
     * @Param
     * @return
     */
    public static void generateExcelToStream(Map<String, Object> params, OutputStream out) {
        // Excel模板文件读取位置
        File xlsTemplatePath = new File("D:\\CodeSpace\\ideaWorkspace\\JXLS-Learning\\src\\main\\resources\\templateFile\\测试报表.xls");

        SimpleDateFormat sdf = new SimpleDateFormat( "MM_dd HH_mm_ss" );
        String timeFormat = sdf.format(new Date());
        String fileName = "temp" + timeFormat +".xls";

        // 生成测试数据
        List<Object> empList = generateDate();

        // 往context中添加数据
        Context context = new Context(params);
        context.putVar("emps", empList);

        try {
            InputStream is = new FileInputStream(xlsTemplatePath);
            JxlsHelper.getInstance().processTemplate(is, out, context);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 读取Excel模板文件，使用jxls渲染，生成报表Excel，存放到指定位置
     * @Author RunningHong
     * @Date 2018/12/4 15:36
     * @Param xlsTemplatePath：模板文件的位置， outPath:存放生成文件的位置
     * @return
     */
    public static void generateExcelToFile(File xlsTemplatePath, File outPath) {
        // 生成临时数据
        List<Object> empList = generateDate();

        try {
            InputStream is = new FileInputStream(xlsTemplatePath);
            OutputStream os = new FileOutputStream(outPath);

            // 往JXLS的Context中添加数据
            Context context = new Context();
            context.putVar("emps", empList);

            // 使用JXLSHelper处理模板
            JxlsHelper.getInstance().processTemplate(is, os, context);
            is.close();
            os.close();

            log.info("Excel模板文件生成报表Excel文件成功！");
            log.info("生成文件路径：" + outPath.getAbsolutePath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 生成Excel2003文件的所有图片到指定位置
     * @Author RunningHong
     * @Date 2018/12/6 10:06
     * @Param
     * @return
     */
    public static void generateExcelPicToFile(HSSFWorkbook workbook, String picSavePath) throws Exception {
        File picSaveFile = new File(picSavePath);
        if (!picSaveFile.exists()) {
            picSaveFile.mkdirs();
        }

        List<HSSFPictureData> pics = workbook.getAllPictures();
        if (pics != null) {
            for (int i = 0; i < pics.size(); i++) {
                HSSFPictureData  picDate = (HSSFPictureData) pics.get(i);
                String saveName = "pic" + (i+1) + new SimpleDateFormat( "_MM_dd HH_mm" ).format(new Date()) + ".jpg";

                FileOutputStream fos = new FileOutputStream(new File(picSavePath + saveName));
                fos.write(picDate.getData());
                fos.close();
            }
        }

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

        String time = new SimpleDateFormat( "(MM_dd HH_mm)" ).format(new Date());
        for (int i = 0; i < 10; i++) {
            Employee emp = new Employee();
            emp.setName("测试人员" + i);
            emp.setAge(i+20);
            emp.setCode(100 + i + "");
            emp.setLinkTel(time);

            list.add(emp);
        }
        return list;
    }


}
