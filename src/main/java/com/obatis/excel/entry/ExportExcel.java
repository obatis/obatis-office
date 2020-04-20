package com.obatis.excel.entry;

import com.obatis.excel.constant.ColumnTypeEnum;
import com.obatis.excel.constant.ExcelConstant;
import com.obatis.excel.param.ExcelParam;
import com.obatis.excel.param.ExcelSubParam;
import com.obatis.tools.ValidateTool;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * excel导出类
 * @author HuangLongPu
 * @date 2018年11月13日10:30:05
 */
public class ExportExcel {

    private static Pattern NUMBER_PATTERN = Pattern.compile("[0-9]+.*[0-9]*");

    /**
     * 设置style
     * @date 2018年12月10日15:21:49
     * @param workbook
     */
    private static HSSFCellStyle getHeaderStyle(HSSFWorkbook workbook){
        HSSFCellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 18);
        headerStyle.setFont(font);
        return headerStyle;
    }

    private static HSSFCellStyle getColumnTitleStyle(HSSFWorkbook workbook){
        HSSFCellStyle columnTitleStyle = workbook.createCellStyle();
        setBorderStyle(columnTitleStyle);
        columnTitleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        columnTitleStyle.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 16);
        columnTitleStyle.setFont(font);
        return columnTitleStyle;
    }

    private static HSSFCellStyle getNormalStyle(HSSFWorkbook workbook){
        HSSFCellStyle dataStyle = workbook.createCellStyle();
        setBorderStyle(dataStyle);
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        dataStyle.setFont(font);
        return dataStyle;
    }

    private static void setBorderStyle(HSSFCellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
    }

    /**
     * 导出
     * @date 2018年11月14日15:34:58
     */
    public static HSSFWorkbook exportExcel(ExcelParam param, List<List<String>> list){

        if(param.getColumn().isEmpty()){
            throw new RuntimeException("导出标题未设置");
        }

        if(list == null || list.isEmpty()) {
            throw new RuntimeException("数据列表为空");
        }

        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet1 = workbook.createSheet("sheet1");

        //标题头开始行
        int headerRowIndex = 0;
        //列标题开始行
        int titleRowIndex = 1;
        int columnNumber = 0;
        List<ExcelSubParam> listHeaderData = new ArrayList<>();

        Map<Integer, Integer> maxWidth = new HashMap<>();
        HSSFRow headerRow = null, columnTitleRow, columnTitleSubRow;
        boolean isExistSubParam = isExistSubParam(param);
        if(ValidateTool.isEmpty(param.getHeader())){
            /**
             * 说明不展示标题，直接从列名开始
             */
            titleRowIndex = 0;
        }else{
            /**
             * 创建标题行
             */
            headerRow = sheet1.createRow(headerRowIndex);
        }

        HSSFCellStyle normalStyle = getNormalStyle(workbook);
        /**
         * 标题下字符串信息描述显示行
          */
        if(!ValidateTool.isEmpty(param.getHeaderMidString())){
            HSSFRow headerMinInfoRow = sheet1.createRow(titleRowIndex);
            headerMinInfoRow.createCell(0).setCellValue(param.getHeaderMidString());
            headerMinInfoRow.getCell(0).setCellStyle(normalStyle);
            titleRowIndex = titleRowIndex + 1;
        }
        if(!ValidateTool.isEmpty(param.getHeaderMidHtml())){
            int rowNumber = setHtmlStyle(param.getHeaderMidHtml(), workbook, titleRowIndex, sheet1);
            titleRowIndex = rowNumber;
        }

        columnTitleRow = sheet1.createRow(titleRowIndex);
        int createDateRow = titleRowIndex;

        /**
         * 创建列名行
         */
        HSSFCellStyle columnTitleStyle = getColumnTitleStyle(workbook);
        if(isExistSubParam){
            createDateRow++;
            /**
             * 创建二级子列名行
             */
            columnTitleSubRow = sheet1.createRow(titleRowIndex + 1);

            createSerial(param.isSerial(), columnTitleStyle, maxWidth, sheet1, columnTitleRow, isExistSubParam, titleRowIndex, columnNumber);
            if(param.isSerial()){ columnNumber++;}
            for(int i=0, j=param.getColumn().size(); i<j; i++){
                List<ExcelSubParam> subParam = param.getColumn().get(i).getSubParam();
                if(subParam.size() > 0){
                    CellRangeAddress region = new CellRangeAddress(titleRowIndex, titleRowIndex, columnNumber, columnNumber + subParam.size() - 1);
                    sheet1.addMergedRegion(region);
                    columnTitleRow.createCell(columnNumber).setCellValue(param.getColumn().get(i).getNameCn());
                    columnTitleRow.getCell(columnNumber).setCellStyle(columnTitleStyle);
                    for(int x =0, y = subParam.size(); x < y; x++){
                        columnTitleSubRow.createCell(columnNumber + x).setCellValue(subParam.get(x).getNameCn());
                        listHeaderData.add(subParam.get(x));
                        setMaxWidth(maxWidth, columnNumber + x, columnTitleSubRow.getCell(columnNumber + x).getStringCellValue().getBytes());
                        columnTitleSubRow.getCell(columnNumber + x);
                    }
                    setMaxWidth(maxWidth, columnNumber, columnTitleRow.getCell(columnNumber).getStringCellValue().getBytes());
                    columnNumber = columnNumber + subParam.size() - 1;
                    columnNumber++;
                }else{
                    CellRangeAddress region = new CellRangeAddress(titleRowIndex, titleRowIndex + 1, columnNumber, columnNumber);
                    sheet1.addMergedRegion(region);
                    columnTitleRow.createCell(columnNumber).setCellStyle(columnTitleStyle);
                    columnTitleRow.getCell(columnNumber).setCellValue(param.getColumn().get(i).getNameCn());
                    listHeaderData.add(param.getColumn().get(i));
                    setMaxWidth(maxWidth, columnNumber, columnTitleRow.getCell(columnNumber).getStringCellValue().getBytes());
                    columnNumber++;
                }
            }
        } else {
            boolean flag = createSerial(param.isSerial(), columnTitleStyle, maxWidth, sheet1, columnTitleRow, isExistSubParam, titleRowIndex, columnNumber);
            int ii=0;
            if(flag){ii = 1;}
            for(int i=0, j=param.getColumn().size(); i<j; i++){
                columnTitleRow.createCell(i + ii).setCellStyle(columnTitleStyle);
                columnTitleRow.getCell(i + ii).setCellValue(param.getColumn().get(i).getNameCn());
                listHeaderData.add(param.getColumn().get(i));
                setMaxWidth(maxWidth, i + ii, columnTitleRow.getCell(i).getStringCellValue().getBytes());
                columnNumber++;
            }
        }

        /**
         * 设置Excel标题
         */
        setHeader(columnNumber, param, sheet1, headerRow, getHeaderStyle(workbook));

        for (int i= 0; i<columnNumber;i++){
            sheet1.setColumnWidth(i, maxWidth.get(i));
        }
        int endRow = setListData(normalStyle,createDateRow + 1, param, listHeaderData, sheet1, list);

        if(null != param.getHeaderEndHtml()){
            setHtmlStyle(param.getHeaderEndHtml(), workbook, endRow, sheet1);
        }

        setMergeOtherCell(columnNumber, param, sheet1, endRow);
        return workbook;
    }


    /**
     * 合并其他单元格
     * @param columnNumber
     * @param param
     * @param sheet1
     */
    private static void setMergeOtherCell(int columnNumber, ExcelParam param, HSSFSheet sheet1, int endRow){
        int upDown = 0;
        if(param.isSerial()){
            columnNumber ++;
        }
        if(null != param.getHeader()){
            upDown = 1;
        }
        if(null != param.getHeaderMidString()){
            CellRangeAddress headerRegion = new CellRangeAddress(upDown, upDown, 0, columnNumber - 1);
            sheet1.addMergedRegion(headerRegion);
            upDown = upDown + 1;
        }
        if(null != param.getHeaderMidHtml()){
            String [] strsp = param.getHeaderMidHtml().split("\n");
            int i = 0;
            for(String s : strsp){
                CellRangeAddress headerRegion = new CellRangeAddress(upDown + i, upDown + i, 0, columnNumber - 1);
                sheet1.addMergedRegion(headerRegion);
                i = i + 1;
            }
        }
        if(null != param.getHeaderEndHtml()){
            String [] strsp = param.getHeaderEndHtml().split("\n");
            int i = 0;
            for(String s : strsp) {
                CellRangeAddress headerRegion = new CellRangeAddress(endRow + i, endRow + i, 0, columnNumber - 1);
                sheet1.addMergedRegion(headerRegion);
                i = i + 1;
            }
        }
    }

    /**
     * 设置标题文件
     * @param columnNumber
     * @param param
     * @param sheet1
     * @param row0
     * @param style
     */
    private static void setHeader(int columnNumber, ExcelParam param, HSSFSheet sheet1, HSSFRow row0, HSSFCellStyle style){
        if(param.isSerial()){columnNumber ++;}
        if(null != row0 && (columnNumber - 1) != 0){
            CellRangeAddress headerRegion = new CellRangeAddress(0, 0, 0, columnNumber - 1);
            sheet1.addMergedRegion(headerRegion);
            row0.createCell(0).setCellValue(param.getHeader());
            row0.getCell(0).setCellStyle(style);
        }
    }

    /**
     * 创建序号
     * @date 2018年12月11日14:34:25
     * @param isExistSubParam
     */
    private static boolean createSerial(boolean isSerial, HSSFCellStyle columnTitleStyle,
                                        Map<Integer, Integer> maxWidth,
                                        HSSFSheet sheet1,
                                        HSSFRow columnTitleRow,
                                        boolean isExistSubParam,
                                        int titleRow,
                                        int columnNumber){
        boolean flag = false;

        if(isSerial){
            if(isExistSubParam){
                CellRangeAddress regionSerial = new CellRangeAddress(titleRow, titleRow + 1, columnNumber, columnNumber);
                sheet1.addMergedRegion(regionSerial);
                columnTitleRow.createCell(columnNumber).setCellValue("序号");
                columnTitleRow.getCell(columnNumber).setCellStyle(columnTitleStyle);
                setMaxWidth(maxWidth, columnNumber, columnTitleRow.getCell(columnNumber).getStringCellValue().getBytes());
                flag = true;
            }else{
                columnTitleRow.createCell(0).setCellStyle(columnTitleStyle);
                columnTitleRow.getCell(0).setCellValue("序号");
                setMaxWidth(maxWidth, 0, columnTitleRow.getCell(0).getStringCellValue().getBytes());
                flag = true;
            }
        }
        return flag;
    }

    /**
     * 根据标题内容设置最大宽度
     * @param maxWidth
     * @param column
     * @param b
     */
    private static void setMaxWidth(Map<Integer,Integer> maxWidth, int column, byte[] b){
        maxWidth.put(column, b.length * 256 + 200);
    }

    /**
     * @date  2018年11月26日14:57:25
     * @param createDataRow
     * @param param
     * @param listHeaderData 二级子元素数据，主要使用样式
     * @param sheet1
     * @param list
     */
    private static int setListData(HSSFCellStyle normalStyle, int createDataRow, ExcelParam param, List<ExcelSubParam> listHeaderData, HSSFSheet sheet1, List<List<String>> list){
        int serial = 0;
        for(int i=0, j=list.size(); i<j; i++){
            HSSFRow row = sheet1.createRow(createDataRow);
            int createOrder = 0;
            int createOrderNow;
            if(param.isSerial()){
                serial++;
                row.createCell(0).setCellValue(serial);
                row.getCell(0).setCellStyle(normalStyle);
                createOrder =  1;
            }
            for(int x=0, y=list.get(i).size(); x<y; x++){

                createOrderNow = createOrder + x;

                ColumnTypeEnum type = listHeaderData.get(x).getType();
                String format = listHeaderData.get(x).getFormat();

                switch (type){
                    case STRING:
                        row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
                        row.getCell(createOrderNow).setCellStyle(normalStyle);
                        break;
                    case NUMBER:
                        Matcher isNum = NUMBER_PATTERN.matcher(list.get(i).get(x));
                        if(isNum.matches()){
                            row.createCell(createOrderNow).setCellStyle(normalStyle);
                            if(ValidateTool.isEmpty(format)){
                                row.getCell(createOrderNow).setCellValue(new BigDecimal(list.get(i).get(x)).toString());
                            }else{
                                row.getCell(createOrderNow).setCellValue(new BigDecimal(list.get(i).get(x)).setScale(Integer.valueOf(format), RoundingMode.HALF_UP).toString());
                            }
                        }else{
                            row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
                            row.getCell(createOrderNow).setCellStyle(normalStyle);
                        }
                        break;
                    case DATE:
                        if(null != list.get(i).get(x)){
                            row.createCell(createOrderNow).setCellValue(ExcelConstant.getDateString(format, list.get(i).get(x)));
                            row.getCell(createOrderNow).setCellStyle(normalStyle);
                        }
                        break;
                    default:
                        row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
                        row.getCell(createOrderNow).setCellStyle(normalStyle);
                }
            }
            createDataRow++;
        }
        return createDataRow;
    }

    /**
     * 是否存在子元素
     * @param param
     * @return
     */
    private static boolean isExistSubParam(ExcelParam param){
        boolean flag = false;
        for(int i=0, j=param.getColumn().size(); i<j; i++){
            if(param.getColumn().get(i).getSubParam().size() > 0){
                flag = true;
                break;
            }
        }
        return flag;
    }

    /**
     * 设置html文本附加格式
     * @param content
     * @param wb
     * @param createRow
     * @param sheet1
     */
    private static int setHtmlStyle(String content, HSSFWorkbook wb, int createRow, HSSFSheet sheet1){

        String[] list = content.split("\n");

        for(String str : list) {
            HSSFRow row = sheet1.createRow(createRow);
            String strSub = "";
            if(str.contains("</strong>")){
                strSub = str.substring(str.indexOf("<strong>") + 8, str.lastIndexOf("</strong>"));
            }
            if(str.contains("</p>")){
                if(!str.contains("</em>") && !str.contains("</span>") && !str.contains("</strong>")){
                    strSub = str.substring(str.indexOf("<p>") + 3, str.lastIndexOf("</p>"));
                }else{
                    if(str.contains("</strong>")){
                        strSub = str.substring(str.indexOf("<strong>") + 8, str.lastIndexOf("</strong>"));
                    }
                    if(str.contains("</em>")){
                        strSub = str.substring(str.indexOf("<em>") + 4, str.lastIndexOf("</em>"));
                    }
                    if(str.contains("</span>")){
                        strSub = str.substring(str.indexOf("<span>") + 6, str.lastIndexOf("</span>"));
                    }
                }
            }
            if(str.contains("</span>")){
                if(str.contains("<span>")){
                    strSub = str.substring(str.indexOf("<span>") + 6, str.lastIndexOf("</span>"));
                }else{
                    strSub = str.substring(str.indexOf("\">") + 2, str.lastIndexOf("</span>"));
                }
            }
            if(str.contains("</em>")){
                strSub = str.substring(str.indexOf("<em>") + 4, str.lastIndexOf("</em>"));
            }
            row.createCell(0).setCellValue(strSub);
            createRow = createRow + 1;

            HSSFCellStyle style = wb.createCellStyle();
            HSSFFont font = wb.createFont();

            if (str.contains("style")){
                String styleString = str.substring(str.indexOf("=") + 2, str.lastIndexOf(";\""));
                String[] styles = styleString.split(":");
                if (styles[0].equals("text-align")) {
                    if (styles[1].trim().equals("center")) {
                        style.setVerticalAlignment(VerticalAlignment.CENTER);
                        style.setAlignment(HorizontalAlignment.CENTER);
                    }
                    if (styles[1].trim().equals("left")) {
                        style.setAlignment(HorizontalAlignment.LEFT);
                    }
                    if (styles[1].trim().equals("right")) {
                        style.setAlignment(HorizontalAlignment.RIGHT);
                    }
                }
                if (styles[0].equals("text-decoration")) {

                    if (styles[1].trim().equals("underline")) {
                        font.setUnderline((byte)1);
                    }
                    if (styles[1].trim().equals("line-through")) {
                        font.setStrikeout(true);
                    }
                }

            }
            if (str.contains("strong")) {
                font.setBold(true);
            }
            if (str.contains("em")) {
                font.setItalic(true);
            }

            style.setFont(font);
            row.getCell(0).setCellStyle(style);
        }

        return createRow;
    }

}
