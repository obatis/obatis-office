package com.obatis.excel.constant;

import com.obatis.convert.date.DefaultDateConstant;

import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * @author HuangLongPu
 * @date 2018年11月27日16:02:45
 */
public class ExcelConstant {

    public static final int TYPE_DATE_FORMAT_DATE = 0;
    public static final int TYPE_DATE_FORMAT_TIME = 1;
    public static final int TYPE_DATE_FORMAT_DAY = 2;
    public static final int TYPE_DATE_FORMAT_YEAR_MONTH = 3;
    public static final int TYPE_DATE_FORMAT_YEAR_MONTH_DAY = 4;

//    public static final int TYPE_NUMBER_FORMAT_TWO_POINT = 2;
//    public static final int TYPE_NUMBER_FORMAT_THREE_POINT = 3;
//    public static final int TYPE_NUMBER_FORMAT_FOUR_POINT = 4;
//    public static final int TYPE_NUMBER_FORMAT_FIVE_POINT = 5;

    /**
     *
     * @param type
     * @param date
     * @return
     */
    public static String getDateString(String type, Object date){

		SimpleDateFormat sdf = null;
        if(type.equals(String.valueOf(TYPE_DATE_FORMAT_DATE))){
        	sdf = DefaultDateConstant.SD_FORMAT_DATE_TIME;
        }else if(type.equals(String.valueOf(TYPE_DATE_FORMAT_TIME))){
            sdf = DefaultDateConstant.SD_FORMAT_TIME_MILLIS;
        }else if(type.equals(String.valueOf(TYPE_DATE_FORMAT_DAY))){
            sdf = DefaultDateConstant.SD_FORMAT_DATE;
        }else if(type.equals(String.valueOf(TYPE_DATE_FORMAT_YEAR_MONTH))){
            sdf = DefaultDateConstant.SD_FORMAT_YEAR_MONTH_JOINT;
        }else if(type.equals(String.valueOf(TYPE_DATE_FORMAT_YEAR_MONTH_DAY))){
            sdf = DefaultDateConstant.SD_FORMAT_DATE_JOINT;
        }

        if(date instanceof Date){
            return sdf.format(date);
        }else{
            return date.toString();
        }
    }


	/**
	 * 表示字符串常规类型
	 */
	public static final int TYPE_FIELD_STRING = 0;
	
	/**
	 * 表示BigDecimal类型
	 */
	public static final int TYPE_FIELD_NUMBER = 1;

	/**
	 * 表示date类型
	 */
	public static final int TYPE_FIELD_DATE = 2;

}
