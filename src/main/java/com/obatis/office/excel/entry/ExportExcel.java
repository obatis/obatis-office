package com.obatis.office.excel.entry;

import com.obatis.office.excel.constant.ColumnTypeEnum;
import com.obatis.office.excel.constant.ExcelConstant;
import com.obatis.office.excel.param.ExcelParam;
import com.obatis.office.excel.param.ExcelSubParam;
import com.obatis.tools.ValidateTool;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * excel导出类
 * @author HuangLongPu
 * @date 2018年11月13日10:30:05
 */
public class ExportExcel {

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

    private static HSSFCellStyle getheaderContextStyle(HSSFWorkbook workbook){
        HSSFCellStyle dataStyle = workbook.createCellStyle();
//        setBorderStyle(dataStyle);
        HSSFFont font = workbook.createFont();
        font.setFontHeightInPoints((short) 14);
        dataStyle.setFont(font);
        return dataStyle;
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
        HSSFSheet sheet = workbook.createSheet("sheet1");

        //标题头开始行
        int headerRowIndex = 0;
        //列标题开始行
        int titleRowIndex = 1;
        List<ExcelSubParam> listHeaderData = new ArrayList<>();

//        Map<Integer, Integer> maxWidth = new HashMap<>();
//        HSSFRow headerRow = null, columnTitleRow, columnTitleSubRow;
        HSSFRow headerRow = null;
        if(ValidateTool.isEmpty(param.getHeader())){
            /**
             * 说明不展示标题，直接从列名开始
             */
            titleRowIndex = 0;
        }else{
            /**
             * 创建标题行
             */
            headerRow = sheet.createRow(headerRowIndex);
        }

        /**
         * 创建列名行
         */
        HSSFCellStyle columnTitleStyle = getColumnTitleStyle(workbook);
        Map<String, Integer> cellInfoMap = new HashMap<>();
        boolean isExistSubParam = isExistSubParam(param.getColumn(), cellInfoMap);
        int subParamChild = cellInfoMap.get("subParamChild");
        int columnTotal = cellInfoMap.get(ExcelConstant.COLUMN_NUMBER_KEY);

        /**
         * 标题下字符串信息描述显示行
          */
        if(!ValidateTool.isEmpty(param.getHeaderMidString())){
            HSSFRow headerMinInfoRow = sheet.createRow(titleRowIndex);
            headerMinInfoRow.createCell(0).setCellValue(param.getHeaderMidString());
            headerMinInfoRow.getCell(0).setCellStyle(getheaderContextStyle(workbook));
            titleRowIndex = titleRowIndex + 1;
        }
        if(!ValidateTool.isEmpty(param.getHeaderMidHtml())){
            int rowNumber = setHtmlStyle(param.getHeaderMidHtml(), workbook, titleRowIndex, sheet);
            titleRowIndex = rowNumber;
        }

        // 创建
        createTitleRow(sheet, cellInfoMap, titleRowIndex, subParamChild);

        HSSFRow columnTitleRow = sheet.getRow(titleRowIndex);
        if(param.isSerial()){
            Integer[] numberArray = getNumberArray(titleRowIndex, titleRowIndex + subParamChild);
            resetRowColIndex(cellInfoMap, 1, numberArray);
        }

        int createDateRow = titleRowIndex;
        if(isExistSubParam){
            createDateRow = titleRowIndex + subParamChild;
            createSerial(param.isSerial(), columnTitleStyle, cellInfoMap, sheet, columnTitleRow, isExistSubParam, titleRowIndex, subParamChild);

            getCellRangeInfo(param.getColumn(), subParamChild, cellInfoMap, ExcelConstant.DEFAULT_INDEX, ExcelConstant.COLUMN_HORIZONTAL_KEY_PREFIX, ExcelConstant.COLUMN_VERTICAL_KEY_PREFIX);
            for(int i = 0, j = param.getColumn().size(); i < j; i++){
                ExcelSubParam excelSubParam = param.getColumn().get(i);
                if (excelSubParam.getSubParam().size() > 0) {
                    createCellRange(param.getColumn().get(i), sheet, columnTitleStyle, listHeaderData, cellInfoMap, titleRowIndex, ExcelConstant.DEFAULT_INDEX, ExcelConstant.COLUMN_HORIZONTAL_KEY_PREFIX, ExcelConstant.COLUMN_VERTICAL_KEY_PREFIX, i, subParamChild);
                } else {
                    int columnNumber = getRowColIndex(cellInfoMap, titleRowIndex);

                    CellRangeAddress region = new CellRangeAddress(titleRowIndex, titleRowIndex + subParamChild, columnNumber, columnNumber);
                    sheet.addMergedRegion(region);

                    setBorderStyle(region, sheet);

//                    HSSFRow columnTitleRow = createRow(sheet, cellInfoMap, titleRowIndex, subParamChild);
                    columnTitleRow.createCell(columnNumber).setCellStyle(columnTitleStyle);
                    columnTitleRow.getCell(columnNumber).setCellValue(excelSubParam.getNameCn());
                    listHeaderData.add(param.getColumn().get(i));
                    setMaxWidth(cellInfoMap, columnNumber, columnTitleRow.getCell(columnNumber).getStringCellValue().getBytes());
                    Integer[] numberArray = getNumberArray(titleRowIndex, titleRowIndex + subParamChild);
                    resetRowColIndex(cellInfoMap, 1, numberArray);
                }
            }
        } else {
//            HSSFRow columnTitleRow = sheet.createRow(titleRowIndex);
            boolean flag = createSerial(param.isSerial(), columnTitleStyle, cellInfoMap, sheet, columnTitleRow, isExistSubParam, titleRowIndex, 0);
            int ii=0;
            if(flag){ii = 1;}
            for(int i=0, j=param.getColumn().size(); i<j; i++){
                columnTitleRow.createCell(i + ii).setCellStyle(columnTitleStyle);
                columnTitleRow.getCell(i + ii).setCellValue(param.getColumn().get(i).getNameCn());
                listHeaderData.add(param.getColumn().get(i));
                setMaxWidth(cellInfoMap, i + ii, columnTitleRow.getCell(i).getStringCellValue().getBytes());
                resetRowColIndex(cellInfoMap, 1, titleRowIndex);
            }
        }

//        int columnNumber = getRowColIndex(cellInfoMap, titleRowIndex);
        /**
         * 设置Excel标题
         */
        setHeader(columnTotal, param, sheet, headerRow, getHeaderStyle(workbook));

//        for (int i= 0; i<columnNumber;i++){
//            sheet.setColumnWidth(i, cellInfoMap.get(i));
//        }
        HSSFCellStyle normalStyle = getNormalStyle(workbook);
        int endRow = setListData(normalStyle,createDateRow + 1, param, listHeaderData, sheet, list);

        if(null != param.getHeaderEndHtml()){
            setHtmlStyle(param.getHeaderEndHtml(), workbook, endRow, sheet);
        }

        setMergeOtherCell(columnTotal, param, sheet, endRow);
        return workbook;
    }

    /**
     * 设置合并单元格的边框
     * @param region
     * @param sheet
     */
    private static void setBorderStyle(CellRangeAddress region, HSSFSheet sheet){
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
    }

    private static Integer[] getNumberArray(int begin, int end) {
        int offset = end - begin;
        Integer[] numberArray = new Integer[offset + 1];
        for(int i = 0; i <= offset; i++) {
            numberArray[i] = begin + i;
        }

        return numberArray;
    }

    private static void createTitleRow(HSSFSheet sheet, Map<String, Integer> cellInfoMap, int titleRowIndex, int subParamChild) {
        for (int i = 0; i <= subParamChild; i++) {
//            if(!cellInfoMap.containsKey(ExcelConstant.ROW_COL_INDEX + "_" + i)) {
//                tempHSSFRow = sheet.createRow(titleRowIndex + i);
//            } else {
//                tempHSSFRow = sheet.getRow(titleRowIndex + i);
//            }

            sheet.createRow(titleRowIndex + i);
            resetRowColIndex(cellInfoMap, 0, i);
        }
    }

    /**
     * 创建单元格
     * @param excelSubParam
     * @param sheet
     * @param handleIndex
     */
    private static void createCellRange(ExcelSubParam excelSubParam, HSSFSheet sheet, HSSFCellStyle columnTitleStyle, List<ExcelSubParam> listHeaderData, Map<String, Integer> cellInfoMap, int titleRowIndex, int index, String columnHorizontalKey, String columnVerticalKey, int handleIndex, int subParamChild) {

        HSSFRow columnTitleRow;
        if(!cellInfoMap.containsKey(ExcelConstant.ROW_COL_INDEX + "_" + (titleRowIndex + index))) {
            columnTitleRow = sheet.createRow(titleRowIndex + (titleRowIndex + index));
        } else {
            columnTitleRow = sheet.getRow(titleRowIndex + index);
            if(columnTitleRow == null) {
                columnTitleRow = sheet.createRow(titleRowIndex + index);
            }
        }

        String tempColumnHorizontalKey = columnHorizontalKey + "_" + handleIndex;
        String tempColumnVerticalKey = columnVerticalKey + "_" + handleIndex;
        int rowColIndex = getRowColIndex(cellInfoMap, titleRowIndex + index);

        int horizontalCell = cellInfoMap.get(tempColumnHorizontalKey);
        int verticalCellCount = cellInfoMap.get(tempColumnVerticalKey);

        if(horizontalCell > 0 || verticalCellCount > 0) {
            // 说明需要合并单元格
            CellRangeAddress region = new CellRangeAddress(titleRowIndex + index, titleRowIndex + index + verticalCellCount, rowColIndex, rowColIndex + horizontalCell);
            sheet.addMergedRegion(region);

            setBorderStyle(region, sheet);
        }

        if(verticalCellCount > 0) {
            Integer[] numberArray = getNumberArray(titleRowIndex + index, titleRowIndex + subParamChild);
            resetRowColIndex(cellInfoMap, horizontalCell + 1, numberArray);
        } else {
            resetRowColIndex(cellInfoMap, horizontalCell + 1, titleRowIndex + index);
        }

        columnTitleRow.createCell(rowColIndex).setCellValue(excelSubParam.getNameCn());
        columnTitleRow.getCell(rowColIndex).setCellStyle(columnTitleStyle);


        List<ExcelSubParam> subParamList = excelSubParam.getSubParam();
        int subParamListSize = subParamList.size();
        if(subParamListSize == 0) {
            setMaxWidth(cellInfoMap, horizontalCell, columnTitleRow.getCell(rowColIndex).getStringCellValue().getBytes());
            listHeaderData.add(excelSubParam);
        }

        for(int i = 0; i < subParamListSize; i++) {
            createCellRange(subParamList.get(i), sheet, columnTitleStyle, listHeaderData, cellInfoMap, titleRowIndex, index + 1, tempColumnHorizontalKey, tempColumnVerticalKey, i, subParamChild);
        }
    }

    private static void resetRowColIndex(Map<String, Integer> cellInfoMap, int addIndex, Integer...indexArray) {
        for (Integer index : indexArray) {
            int rowColIndex = getRowColIndex(cellInfoMap, index);
            cellInfoMap.put(ExcelConstant.ROW_COL_INDEX + "_" + index, rowColIndex + addIndex);
        }
    }

    private static int getRowColIndex(Map<String, Integer> cellInfoMap, int index) {
        Integer rowColIndex = cellInfoMap.get(ExcelConstant.ROW_COL_INDEX + "_" + index);
        if(rowColIndex == null) {
            rowColIndex = 0;
        }
        return rowColIndex;
    }

    /**
     * 处理单元格坐标占位
     * @param column
     * @param subParamChild
     * @param cellInfoMap
     * @param index
     * @param columnHorizontalKey
     * @param columnVerticalKey
     */
    private static void getCellRangeInfo(List<ExcelSubParam> column, int subParamChild, Map<String, Integer> cellInfoMap, int index, String columnHorizontalKey, String columnVerticalKey) {
        for (int i = 0, j = column.size(); i < j; i++) {
            String tempColumnHorizontalKey = columnHorizontalKey + "_" + i;
            String tempColumnVerticalKey = columnVerticalKey + "_" + i;
            ExcelSubParam subParam = column.get(i);
            List<ExcelSubParam> subParamArray = subParam.getSubParam();
            int subParamArraySize = subParamArray.size();
            if(subParamArraySize > 0) {
                cellInfoMap.put(tempColumnHorizontalKey, subParamArraySize - 1);
                cellInfoMap.put(tempColumnVerticalKey, 0);
                 // 重置上级单元格的横坐标跨度
                resetCellRangeInfo(cellInfoMap, columnHorizontalKey, subParamArraySize);
                getCellRangeInfo(subParamArray, subParamChild, cellInfoMap, index + 1, tempColumnHorizontalKey, tempColumnVerticalKey);
            } else {
                cellInfoMap.put(tempColumnHorizontalKey, 0);
                cellInfoMap.put(tempColumnVerticalKey, subParamChild - index);
            }
        }
    }

    /**
     * 重置上级单元格的横坐标跨度
     * @param cellInfoMap
     * @param columnHorizontalKey
     * @param addColumn
     */
    private static void resetCellRangeInfo(Map<String, Integer> cellInfoMap, String columnHorizontalKey, int addColumn) {
        if(ExcelConstant.COLUMN_HORIZONTAL_KEY_PREFIX.equals(columnHorizontalKey)) {
            return;
        }

        cellInfoMap.put(columnHorizontalKey, cellInfoMap.get(columnHorizontalKey) + (addColumn - 1));
        resetCellRangeInfo(cellInfoMap, columnHorizontalKey.substring(0, columnHorizontalKey.lastIndexOf("_")), addColumn);
    }

    /**
     * 合并其他单元格
     * @param columnNumber
     * @param param
     * @param sheet1
     */
    private static void setMergeOtherCell(int columnNumber, ExcelParam param, HSSFSheet sheet1, int endRow){
        int upDown = 0;
//        if(param.isSerial()){
//            columnNumber ++;
//        }
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
     * @param sheet
     * @param headerRow
     * @param style
     */
    private static void setHeader(int columnNumber, ExcelParam param, HSSFSheet sheet, HSSFRow headerRow, HSSFCellStyle style){
        if(null != headerRow && (columnNumber - 1) != 0){
            CellRangeAddress headerRegion = new CellRangeAddress(0, 0, 0, columnNumber - 1);
            sheet.addMergedRegion(headerRegion);

//            setBorderStyle(headerRegion, sheet);

            headerRow.createCell(0).setCellValue(param.getHeader());
            headerRow.getCell(0).setCellStyle(style);
        }
    }

    /**
     * 创建序号
     * @date 2018年12月11日14:34:25
     * @param isExistSubParam
     */
    private static boolean createSerial(boolean isSerial, HSSFCellStyle columnTitleStyle,
                                        Map<String, Integer> maxWidth,
                                        HSSFSheet sheet,
                                        HSSFRow columnTitleRow,
                                        boolean isExistSubParam,
                                        int titleRow, int subParamChild){
        boolean flag = false;

//        HSSFRow columnTitleRow = sheet.getRow(titleRow);
//        if(columnTitleRow == null) {
//            columnTitleRow = sheet.createRow(titleRow);
//
//        }

        if(isSerial){
            if(isExistSubParam){
                CellRangeAddress regionSerial = new CellRangeAddress(titleRow, titleRow + subParamChild, 0, 0);
                sheet.addMergedRegion(regionSerial);
                columnTitleRow.createCell(0).setCellValue("序号");
                columnTitleRow.getCell(0).setCellStyle(columnTitleStyle);
                setMaxWidth(maxWidth, 0, columnTitleRow.getCell(0).getStringCellValue().getBytes());
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
    private static void setMaxWidth(Map<String,Integer> maxWidth, int column, byte[] b){
        maxWidth.put(column + "", b.length * 256 + 200);
    }

    /**
     * @date  2018年11月26日14:57:25
     * @param createDataRow
     * @param param
     * @param listHeaderData 二级子元素数据，主要使用样式
     * @param sheet
     * @param list
     */
    private static int setListData(HSSFCellStyle normalStyle, int createDataRow, ExcelParam param, List<ExcelSubParam> listHeaderData, HSSFSheet sheet, List<List<String>> list){
        int serial = 0;
        for(int i=0, j=list.size(); i<j; i++){
            HSSFRow row = sheet.createRow(createDataRow);
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
//                        Matcher isNum = NUMBER_PATTERN.matcher(list.get(i).get(x));
                        if(ValidateTool.isNumber(list.get(i).get(x))){
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
     * @param column
     * @return
     */
    private static boolean isExistSubParam(List<ExcelSubParam> column, Map<String, Integer> titleParamInfoMap){
        int subParamChild = 0;

        for (int i = 0, j = column.size(); i < j; i++) {
            ExcelSubParam subParam = column.get(i);
            int tempSubParamChild = getSubParamChildInfo(subParam, subParamChild, titleParamInfoMap);
            if(tempSubParamChild > subParamChild) {
                subParamChild = tempSubParamChild;
            }
        }

        titleParamInfoMap.put("subParamChild", subParamChild);
        return subParamChild > 0;
    }

    private static int getSubParamChildInfo(ExcelSubParam subParam, int subParamChild, Map<String, Integer> titleParamInfoMap) {

        if(subParam.getSubParam().isEmpty()) {
            Integer columnNumber = titleParamInfoMap.get(ExcelConstant.COLUMN_NUMBER_KEY);
            if(columnNumber == null) {
                columnNumber = 0;
            }
            columnNumber ++;
            titleParamInfoMap.put(ExcelConstant.COLUMN_NUMBER_KEY, columnNumber);
            return subParamChild;
        }

        subParamChild ++;
        int maxTempParamChild = 0;

        for (ExcelSubParam subParamItem : subParam.getSubParam()) {
            int tempSubParamChild = getSubParamChildInfo(subParamItem, subParamChild, titleParamInfoMap);
            if(tempSubParamChild > maxTempParamChild) {
                maxTempParamChild = tempSubParamChild;
            }
        }
        return maxTempParamChild;
    }

    /**
     * 设置html文本附加格式
     * @param content
     * @param workbook
     * @param createRow
     * @param sheet
     */
    private static int setHtmlStyle(String content, HSSFWorkbook workbook, int createRow, HSSFSheet sheet){

        String[] list = content.split("\n");

        for(String str : list) {
            HSSFRow row = sheet.createRow(createRow);
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

            HSSFCellStyle style = workbook.createCellStyle();
            HSSFFont font = workbook.createFont();

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
