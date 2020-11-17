package com.obatis.office.excel.constant;

import com.obatis.convert.date.DateConvert;
import com.obatis.tools.ValidateTool;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Date;

/**
 * @author HuangLongPu
 * @date 2018年11月27日16:02:45
 */
public class ExcelConstant {

    public static final String COLUMN_HORIZONTAL_KEY_PREFIX = "ch";
    public static final String COLUMN_VERTICAL_KEY_PREFIX = "cv";

    public static final String ROW_COL_INDEX = "row";
    public static final String COLUMN_NUMBER_KEY = "column_number";

    public static final int DEFAULT_INDEX = 0;

    /**
     * @param format
     * @param date
     * @return
     */
    public static String getDateString(String format, Object date){

		if(ValidateTool.isEmpty(date)) {
		    return null;
        } else if(date instanceof Date) {
            LocalDateTime dateTime = DateConvert.parseDateTimeByMilli(((Date) date).getTime());
            if(ValidateTool.isEmpty(format)) {
                return DateConvert.formatDateTime(dateTime);
            } else {
		        return DateConvert.formatDateTime(dateTime, format);
            }
        } else if (date instanceof LocalDate) {
		    if(ValidateTool.isEmpty(format)) {
		        return DateConvert.formatDate((LocalDate) date);
            } else {
		        return DateConvert.formatDate((LocalDate) date, format);
            }
        } else if (date instanceof LocalDateTime) {
            if(ValidateTool.isEmpty(format)) {
                return DateConvert.formatDateTime((LocalDateTime) date);
            } else {
                return DateConvert.formatDate((LocalDate) date, format);
            }
        } else {
		    return date.toString();
        }

    }

}
