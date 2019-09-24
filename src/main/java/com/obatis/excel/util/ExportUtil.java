package com.obatis.excel.util;

import com.obatis.excel.constant.ExcelConstant;
import com.obatis.excel.param.ExcelParam;
import com.obatis.excel.param.ExcelSubParam;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileOutputStream;
import java.io.IOException;
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
 * @author heChengBo
 * @date 2018年11月13日10:30:05
 */
public class ExportUtil {

    private static Pattern NUMBER_PATTERN = Pattern.compile("[0-9]+.*[0-9]*");

    /**
     * web端调用直接导出文档类
     * @param response web传入
     * @param param 标题参数传入
     * @param list 导出数据
     * @param fileName 导出文件名称
     */
//    public static void exportExcelDownload(HttpServletResponse response, ExcelParam param, List<List<String>> list, String fileName){
//        try {
//            HSSFWorkbook wb = ExportUtil.exportExcel(param, list);
//            response.setContentType("application/vnd.ms-excel");
//            response.setHeader("Content-Disposition", "attachment;filename=" + URLEncoder.encode(fileName + ".xls", "utf-8"));
//            OutputStream outputStream = response.getOutputStream();
//            wb.write(outputStream);
//            outputStream.flush();
//            outputStream.close();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }


    /**
     * 设置style
     * @date 2018年12月10日15:21:49
     * @param wb
     */
    private static HSSFCellStyle getHeaderStyle(HSSFWorkbook wb){
        HSSFCellStyle style = wb.createCellStyle();
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = wb.createFont();
        font.setBold(true);
        style.setFont(font);
        return style;
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

        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet1 = wb.createSheet("Sheet1");
        HSSFCellStyle style = getHeaderStyle(wb);
        //标题头开始行
        int headerRow = 0;
        //列标题开始行
        int titleRow = 1;
        int columnNumber = 0;
        List<ExcelSubParam> listHeaderData = new ArrayList<>();

        Map<Integer, Integer> maxWidth = new HashMap<>(16);
        HSSFRow row0 = null, row1 = null, row2 = null;
        boolean isExistSubParam = isExistSubParam(param);
        if(null == param.getHeader()){
            titleRow = 0;
        }else{
            row0 = sheet1.createRow(headerRow);
        }

        if(null != param.getHeaderMidString()){
            HSSFRow row = sheet1.createRow(titleRow);
            row.createCell(0).setCellValue(param.getHeaderMidString());
            titleRow = titleRow + 1;

        }
        if(null != param.getHeaderMidHtml()){
//            HSSFRow row = sheet1.createRow(titleRow);
//            row.createCell(0).setCellValue(param.getHeaderMidHtml());
            int rowNumber = setHtmlStyle(param.getHeaderMidHtml(), wb, titleRow, sheet1);
//            titleRow = titleRow + 1;
            titleRow = rowNumber;
        }

        row1 = sheet1.createRow(titleRow);
        int createDateRow = titleRow;

        if(isExistSubParam){
            createDateRow++;
            row2 = sheet1.createRow(titleRow + 1);

            createSerial(param.isSerial(), maxWidth, sheet1, row1, style, isExistSubParam, titleRow, columnNumber);
            if(param.isSerial()){ columnNumber++;}
            for(int i=0, j=param.getColumn().size(); i<j; i++){
                List<ExcelSubParam> subParam = param.getColumn().get(i).getSubParam();
                if(subParam.size() > 0){
                    CellRangeAddress region = new CellRangeAddress(titleRow, titleRow, columnNumber, columnNumber + subParam.size() - 1);
                    sheet1.addMergedRegion(region);
                    row1.createCell(columnNumber).setCellValue(param.getColumn().get(i).getNameCn());
                    row1.getCell(columnNumber).setCellStyle(style);
                    for(int x =0, y=subParam.size(); x<y; x++){
                        row2.createCell(columnNumber + x).setCellValue(subParam.get(x).getNameCn());
                        listHeaderData.add(subParam.get(x));
                        setMaxWidth(maxWidth, columnNumber + x, row2.getCell(columnNumber + x).getStringCellValue().getBytes());
                        row2.getCell(columnNumber + x).setCellStyle(style);
                    }
                    setMaxWidth(maxWidth, columnNumber, row1.getCell(columnNumber).getStringCellValue().getBytes());
                    columnNumber = columnNumber + subParam.size() - 1;
                    columnNumber++;
                }else{
                    CellRangeAddress region = new CellRangeAddress(titleRow, titleRow + 1, columnNumber, columnNumber);
                    sheet1.addMergedRegion(region);
                    row1.createCell(columnNumber).setCellStyle(style);
                    row1.getCell(columnNumber).setCellValue(param.getColumn().get(i).getNameCn());
                    listHeaderData.add(param.getColumn().get(i));
                    setMaxWidth(maxWidth, columnNumber, row1.getCell(columnNumber).getStringCellValue().getBytes());
                    columnNumber++;
                }
            }
        }else{
            boolean flag = createSerial(param.isSerial(), maxWidth, sheet1, row1, style, isExistSubParam, titleRow, columnNumber);
            int ii=0;
            if(flag){ii = 1;}
            for(int i=0, j=param.getColumn().size(); i<j; i++){
                row1.createCell(i + ii).setCellStyle(style);
                row1.getCell(i + ii).setCellValue(param.getColumn().get(i).getNameCn());
                listHeaderData.add(param.getColumn().get(i));
                setMaxWidth(maxWidth, i + ii, row1.getCell(i).getStringCellValue().getBytes());
                columnNumber++;
            }
        }
        setHeader(columnNumber, param, sheet1, row0, style);

        for (int i= 0; i<columnNumber;i++){
            sheet1.setColumnWidth(i, maxWidth.get(i));
        }
        int endRow = setListData(createDateRow + 1, param, listHeaderData, sheet1, list);

        if(null != param.getHeaderEndHtml()){
            setHtmlStyle(param.getHeaderEndHtml(), wb, endRow, sheet1);
        }
        setMergeOtherCell(columnNumber, param, sheet1, endRow);

        return wb;
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
    private static boolean createSerial(boolean isSerial,
                                        Map<Integer, Integer> maxWidth,
                                        HSSFSheet sheet1,
                                        HSSFRow row1,
                                        HSSFCellStyle style,
                                        boolean isExistSubParam,
                                        int titleRow,
                                        int columnNumber){
        boolean flag = false;

        if(isSerial){
            if(isExistSubParam){
                CellRangeAddress regionSerial = new CellRangeAddress(titleRow, titleRow + 1, columnNumber, columnNumber);
                sheet1.addMergedRegion(regionSerial);
                row1.createCell(columnNumber).setCellValue("序号");
                setMaxWidth(maxWidth, columnNumber, row1.getCell(columnNumber).getStringCellValue().getBytes());
                flag = true;
            }else{
                row1.createCell(0).setCellStyle(style);
                row1.getCell(0).setCellValue("序号");
                setMaxWidth(maxWidth, 0, row1.getCell(0).getStringCellValue().getBytes());
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
    private static int setListData(int createDataRow, ExcelParam param, List<ExcelSubParam> listHeaderData, HSSFSheet sheet1, List<List<String>> list){
        int serial = 0;
        for(int i=0, j=list.size(); i<j; i++){
            HSSFRow row = sheet1.createRow(createDataRow);
            int createOrder = 0;
            int createOrderNow;
            if(param.isSerial()){
                serial++;
                row.createCell(0).setCellValue(serial);
                createOrder =  1;
            }
            for(int x=0, y=list.get(i).size(); x<y; x++){

                createOrderNow = createOrder + x;

                String type = listHeaderData.get(x).getType();
                String format = listHeaderData.get(x).getFormat();

                switch (Integer.valueOf(type)){
                    case ExcelConstant.TYPE_FIELD_STRING:
                        row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
                        break;
                    case ExcelConstant.TYPE_FIELD_NUMBER:
                        Matcher isNum = NUMBER_PATTERN.matcher(list.get(i).get(x));
                        if(isNum.matches()){
                            if(null == format){
                                row.createCell(createOrderNow).setCellValue(new BigDecimal(list.get(i).get(x)).toString());
                            }else{
                                row.createCell(createOrderNow).setCellValue(new BigDecimal(list.get(i).get(x)).setScale(Integer.valueOf(format), RoundingMode.HALF_UP).toString());
                            }
                        }else{
                            row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
                        }
                        break;
                    case ExcelConstant.TYPE_FIELD_DATE:
                        if(null != list.get(i).get(x)){
                            row.createCell(createOrderNow).setCellValue(ExcelConstant.getDateString(type, list.get(i).get(x)));
                        }
                        break;
                    default:
                        row.createCell(createOrderNow).setCellValue(list.get(i).get(x));
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



//        String split = "<p style=\"text-align: center;\"><strong>本报表是由XXX提供</strong></p>\n" +
//                "<p style=\"text-align: center;\"><strong>本报表是由XXX提供</strong></p>";

        String[] list = content.split("\n");


//        style.setVerticalAlignment(VerticalAlignment.CENTER);
//        style.setAlignment(HorizontalAlignment.CENTER);
//        HSSFFont font = wb.createFont();
//        font.setBold(true);
//        style.setFont(font);


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


    public static void main(String args[]){

//        String split = "<p><strong>本报表是由XXX提供</strong></p>";
//
//        //去回车
//        //split = split.replaceAll("\r|\n", "");
//
//        String[] list = split.split("\n");
//
//        for(String s : list){
//            String style = s.substring(s.indexOf("=")+2, s.lastIndexOf(";\""));
//
//            String[] styles = style.split(":");
//
//            if(styles[0].equals("text-align")){
//                if(styles[1].trim().equals("center")){
//                    System.out.println("test");
//                }
//                if(styles[1].trim().equals("left")){
//                    System.out.println("test");
//                }
//                if(styles[1].trim().equals("right")){
//                    System.out.println("test");
//                }
//            }
//            if(styles[0].equals("text-decoration")){
//                if(styles[1].trim().equals("underline")){
//
//                }
//                if(styles[1].trim().equals("line-through")){
//
//                }
//            }
//
//            if(s.contains("strong")){
//                System.out.println("strong");
//            }
//            if(s.contains("em")){
//                System.out.println("strong");
//            }
//
//            System.out.println(s);
//        }



        try {

            ExcelParam param = new ExcelParam();
            param.setHeaderMidString("创建时间：2019-1-24 10:30:40  创建人：张三   交易状态：交易中 ");
//            param.setHeaderMidHtml("<p style=\"text-align: center;\"><strong>本报表是由XXX提供</strong></p>\n" +
//                    "<p style=\"text-align: center;\"><strong>本报表是由XXX提供</strong></p>");
//            param.setHeaderEndHtml("<p><strong>本报表由作者上传并发布。</strong></p>\n" +
//                    "<p><em>文章仅代表作者个人观点，不代表立场。</em></p>\n" +
//                    "<p><span style=\"text-decoration: underline;\">未经作者许可，不得转载。</span></p>");

            param.setHeaderMidHtml("<p>这个是表头</p>");
//            param.setHeaderEndHtml("<p><strong>本报表由作者上传并发布。</strong></p>");
param.setHeaderEndHtml("<p><span style=\"text-decoration: underline;\">这个是表尾</span></p>");
            param.setHeader("测试文档");

            List<ExcelSubParam> list = new ArrayList<ExcelSubParam>();
            list.add(new ExcelSubParam());
            list.add(new ExcelSubParam());
            list.add(new ExcelSubParam());
            list.add(new ExcelSubParam());
            list.add(new ExcelSubParam());
            list.add(new ExcelSubParam());

            list.get(0).setNameCn("姓名1");
            list.get(0).setType("1");
            list.get(0).setFormat("2");

            list.get(1).setNameCn("资金2");
            list.get(1).setType("1");
            list.get(1).setFormat("2");

            List<ExcelSubParam> listSub = new ArrayList<ExcelSubParam>();
            listSub.add(new ExcelSubParam());
            listSub.add(new ExcelSubParam());

            listSub.get(0).setNameCn("余额2-1");
            listSub.get(0).setType("1");
            listSub.get(0).setFormat("2");

            listSub.get(1).setNameCn("银行余额2-2");
            listSub.get(1).setType("1");
            listSub.get(1).setFormat("2");

//            list.get(1).setSubParam(listSub);


            list.get(2).setNameCn("第三个3");
            list.get(2).setFormat("2");

            list.get(3).setNameCn("第四个");
            list.get(3).setFormat("2");


            list.get(4).setNameCn("资金5");
            list.get(4).setType("1");
            list.get(4).setFormat("2");

            List<ExcelSubParam> listSub4 = new ArrayList<ExcelSubParam>();
            listSub4.add(new ExcelSubParam());
            listSub4.add(new ExcelSubParam());
            listSub4.add(new ExcelSubParam());

            listSub4.get(0).setNameCn("余额5_1");
            listSub4.get(0).setFormat("2");

            listSub4.get(1).setNameCn("银行余额5_2");
            listSub4.get(1).setType("1");
            listSub4.get(1).setFormat("2");

            listSub4.get(2).setNameCn("期初余额5_2");
            listSub4.get(2).setFormat("2");

//            list.get(4).setSubParam(listSub4);

            list.get(5).setNameCn("第22个");
            list.get(5).setType("1");
            list.get(5).setFormat("2");


            param.setColumn(list);
//            param.setSerial(true);
//
            List<List<String>> listTest = new ArrayList<List<String>>();
            List<String> listStr1 = new ArrayList<String>();
            List<String> listStr2 = new ArrayList<String>();
            String s1 = "test";
            String s2 = "bobo";
            listStr1.add(s1);
            listStr1.add("23.919");
            listStr1.add("23.919");
            listStr1.add("23.919");
            listStr1.add("23.919");
            listStr1.add("23.919");

            listStr2.add(s2);
            listStr2.add("288.919");
            listTest.add(listStr1);
            listTest.add(listStr2);

            HSSFWorkbook wb = exportExcel(param, listTest);

            FileOutputStream fout = new FileOutputStream("D:/Temp/test.xls");
            wb.write(fout);
            fout.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
