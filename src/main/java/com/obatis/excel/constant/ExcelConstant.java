package com.obatis.excel.constant;

import com.obatis.convert.date.DefaultDateConstant;
import com.obatis.tools.ValidateTool;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author HuangLongPu
 * @date 2018年11月27日16:02:45
 */
public class ExcelConstant {

    /**
     * @param format
     * @param date
     * @return
     */
    public static String getDateString(String format, Object date){

		SimpleDateFormat sdf = null;

		if(ValidateTool.isEmpty(date)) {
		    return null;
        } else if(ValidateTool.isEmpty(format) || DefaultDateConstant.DATE_TIME_PATTERN.equals(format)) {
            sdf = DefaultDateConstant.SD_FORMAT_DATE_TIME;
        } else if (DefaultDateConstant.TIME_MILLIS_PATTERN.equals(format)) {
            sdf = DefaultDateConstant.SD_FORMAT_TIME_MILLIS;
        } else if (DefaultDateConstant.DATE_PATTERN.equals(format)) {
            sdf = DefaultDateConstant.SD_FORMAT_DATE;
        } else if (DefaultDateConstant.YEAR_MONTH_JOINT_PATTERN.equals(format)) {
            sdf = DefaultDateConstant.SD_FORMAT_YEAR_MONTH_JOINT;
        } else if (DefaultDateConstant.DATE_JOINT_PATTERN.equals(format)) {
            sdf = DefaultDateConstant.SD_FORMAT_DATE_JOINT;
        }

        if(date instanceof Date){
            return sdf.format(date);
        }else{
            return date.toString();
        }
    }

}
