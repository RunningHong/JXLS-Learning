package com.zh.jxls.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.jboss.logging.Param;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;

/**
 * 将Excel版本(2007+)转换为2003版
 * @author RunningHong
 * @create 2018-12-03 21:47
 */
public class ExcelXlsxToXlsUtil {

   /**
     * 入口方法<br>
     * 将Excel文件转换为Excle2003版本的HSSFWorkbook
     * @Author RunningHong
     * @Date 2018/12/4 10:35
     * @param transExcelFile 需要转化的文件
     * @return Excel的HSSFWorkbook
     */
    public HSSFWorkbook transformExcleTo2003(File transExcelFile) {
        HSSFWorkbook excelBook = new HSSFWorkbook();
        try {
            InputStream input = new FileInputStream(transExcelFile);

            // 获取文件的类型(扩展名)
            String fileType = FileUtil.getFileType(transExcelFile);
            if(".xls".equals(fileType)) { // Excel类型为2003
                excelBook = new HSSFWorkbook(input);
            } else if(".xlsx".equals(fileType)) { // Excel类型为2007+
                XSSFWorkbook workbook2007 = new XSSFWorkbook(input);
                transformXSSF(workbook2007, excelBook);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            return excelBook;
        }
    }


    private int lastColumn = 0;
    private HashMap<Integer, HSSFCellStyle> styleMap = new HashMap();

    /**
     * XSSFSheet是Excel2003的格式，HSSFSheet是Excel2007+的格式
     * 此处遍历sheet，进行sheet内容替换
     * @Author RunningHong
     * @Date 2018/12/4 9:44
     * @Param
     * @return
     */
    public void transformXSSF(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew) {
        HSSFSheet sheetNew;
        XSSFSheet sheetOld;

        workbookNew.setMissingCellPolicy(workbookOld.getMissingCellPolicy());

        for (int i = 0; i < workbookOld.getNumberOfSheets(); i++) {
            sheetOld = workbookOld.getSheetAt(i);
            sheetNew = workbookNew.getSheet(sheetOld.getSheetName());
            sheetNew = workbookNew.createSheet(sheetOld.getSheetName());
            this.transformSheet(workbookOld, workbookNew, sheetOld, sheetNew);
        }
    }

    /**
     * 对sheet内容进行转换
     * @Author RunningHong
     * @Date 2018/12/4 9:46
     * @Param
     * @return
     */
    private void transformSheet(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew,
                                XSSFSheet sheetOld, HSSFSheet sheetNew) {

        sheetNew.setDisplayFormulas(sheetOld.isDisplayFormulas());
        sheetNew.setDisplayGridlines(sheetOld.isDisplayGridlines());
        sheetNew.setDisplayGuts(sheetOld.getDisplayGuts());
        sheetNew.setDisplayRowColHeadings(sheetOld.isDisplayRowColHeadings());
        sheetNew.setDisplayZeros(sheetOld.isDisplayZeros());
        sheetNew.setFitToPage(sheetOld.getFitToPage());

        sheetNew.setHorizontallyCenter(sheetOld.getHorizontallyCenter());
        sheetNew.setMargin(Sheet.BottomMargin, sheetOld.getMargin(Sheet.BottomMargin));
        sheetNew.setMargin(Sheet.FooterMargin, sheetOld.getMargin(Sheet.FooterMargin));
        sheetNew.setMargin(Sheet.HeaderMargin, sheetOld.getMargin(Sheet.HeaderMargin));
        sheetNew.setMargin(Sheet.LeftMargin, sheetOld.getMargin(Sheet.LeftMargin));
        sheetNew.setMargin(Sheet.RightMargin, sheetOld.getMargin(Sheet.RightMargin));
        sheetNew.setMargin(Sheet.TopMargin, sheetOld.getMargin(Sheet.TopMargin));
        sheetNew.setPrintGridlines(sheetNew.isPrintGridlines());
        sheetNew.setRightToLeft(sheetNew.isRightToLeft());
        sheetNew.setRowSumsBelow(sheetNew.getRowSumsBelow());
        sheetNew.setRowSumsRight(sheetNew.getRowSumsRight());
        sheetNew.setVerticallyCenter(sheetOld.getVerticallyCenter());

        HSSFRow rowNew;
        for (Row row : sheetOld) {
            rowNew = sheetNew.createRow(row.getRowNum());
            if (rowNew != null)
                this.transformRow(workbookOld, workbookNew, (XSSFRow) row, rowNew);
        }

        for (int i = 0; i < this.lastColumn; i++) {
            sheetNew.setColumnWidth(i, sheetOld.getColumnWidth(i));
            sheetNew.setColumnHidden(i, sheetOld.isColumnHidden(i));
        }

        for (int i = 0; i < sheetOld.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheetOld.getMergedRegion(i);
            sheetNew.addMergedRegion(merged);
        }
    }

    /**
     * 对行元素row进行转换
     * @Author RunningHong
     * @Date 2018/12/4 9:47
     * @Param
     * @return
     */
    private void transformRow(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew,
                           XSSFRow rowOld, HSSFRow rowNew) {
        HSSFCell cellNew;
        rowNew.setHeight(rowOld.getHeight());

        for (Cell cell : rowOld) {
            cellNew = rowNew.createCell(cell.getColumnIndex(), cell.getCellType());
            if (cellNew != null) {
                this.transformCell(workbookOld, workbookNew, (XSSFCell) cell, cellNew);
            }
        }

        this.lastColumn = Math.max(this.lastColumn, rowOld.getLastCellNum());
    }

    /**
     * 对行内元素进行cell转换
     * @Author RunningHong
     * @Date 2018/12/4 9:49
     * @Param
     * @return
     */
    private void transformCell(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew,
                           XSSFCell cellOld, HSSFCell cellNew) {
        cellNew.setCellComment(cellOld.getCellComment());

        Integer hash = cellOld.getCellStyle().hashCode();
        if (this.styleMap != null && !this.styleMap.containsKey(hash)) {
            this.transformStyle(workbookOld, workbookNew, hash,
                           cellOld.getCellStyle(),
                           (HSSFCellStyle) workbookNew.createCellStyle());
        }
        cellNew.setCellStyle(this.styleMap.get(hash));

        cellOld.getCellType();
        switch (cellOld.getCellTypeEnum()) {
            case BLANK:
                break;
            case BOOLEAN:
                cellNew.setCellValue(cellOld.getBooleanCellValue());
                break;
            case ERROR:
                cellNew.setCellValue(cellOld.getErrorCellValue());
                break;
            case FORMULA:
                cellNew.setCellValue(cellOld.getCellFormula());
                break;
            case NUMERIC:
                cellNew.setCellValue(cellOld.getNumericCellValue());
                break;
            case STRING:
                cellNew.setCellValue(cellOld.getStringCellValue());
                break;
            default:
                System.out.println("transform: Unbekannter Zellentyp "
                        + cellOld.getCellType());
        }
    }

    /**
     * 对单元格cell的样式进行转化
     * @Author RunningHong
     * @Date 2018/12/4 9:51
     * @Param
     * @return
     */
    private void transformStyle(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew,
                                Integer hash, XSSFCellStyle styleOld, HSSFCellStyle styleNew) {

        styleNew.setAlignment(styleOld.getAlignment());
        styleNew.setBorderBottom(styleOld.getBorderBottom());
        styleNew.setBorderLeft(styleOld.getBorderLeft());
        styleNew.setBorderRight(styleOld.getBorderRight());
        styleNew.setBorderTop(styleOld.getBorderTop());
        styleNew.setDataFormat(this.transformDataFormat(workbookOld, workbookNew, styleOld.getDataFormat()));
        styleNew.setFillBackgroundColor(styleOld.getFillBackgroundColor());
        styleNew.setFillForegroundColor(styleOld.getFillForegroundColor());
        styleNew.setFillPattern(styleOld.getFillPattern());
        styleNew.setFont(this.transformFont(workbookNew, (XSSFFont) styleOld.getFont()));
        styleNew.setHidden(styleOld.getHidden());
        styleNew.setIndention(styleOld.getIndention());
        styleNew.setLocked(styleOld.getLocked());
        styleNew.setVerticalAlignment(styleOld.getVerticalAlignment());
        styleNew.setWrapText(styleOld.getWrapText());
        this.styleMap.put(hash, styleNew);
    }

    /**
     * 数据格式机型转化
     * @Author RunningHong
     * @Date 2018/12/4 9:53
     * @Param
     * @return
     */
    private short transformDataFormat(XSSFWorkbook workbookOld, HSSFWorkbook workbookNew, short index) {
        DataFormat formatOld = workbookOld.createDataFormat();
        DataFormat formatNew = workbookNew.createDataFormat();
        return formatNew.getFormat(formatOld.getFormat(index));
    }

    /**
     * 字体进行转化
     * @Author RunningHong
     * @Date 2018/12/4 9:53
     * @Param
     * @return
     */
    private HSSFFont transformFont(HSSFWorkbook workbookNew, XSSFFont fontOld) {
        HSSFFont fontNew = workbookNew.createFont();
        fontNew.setBold(fontOld.getBold());
        fontNew.setCharSet(fontOld.getCharSet());
        fontNew.setColor(fontOld.getColor());
        fontNew.setFontName(fontOld.getFontName());
        fontNew.setFontHeight(fontOld.getFontHeight());
        fontNew.setItalic(fontOld.getItalic());
        fontNew.setStrikeout(fontOld.getStrikeout());
        fontNew.setTypeOffset(fontOld.getTypeOffset());
        fontNew.setUnderline(fontOld.getUnderline());
        return fontNew;
    }
}
