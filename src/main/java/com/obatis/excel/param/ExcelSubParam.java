package com.obatis.excel.param;

import java.util.ArrayList;
import java.util.List;

/**
 * @author HuangLongPu
 * @date 2018年11月14日10:33:29
 */
public class ExcelSubParam {

    /**
     * 字段显示名称
     */
    private String nameCn;
    /**
     * 字段类型
     */
    private String type = "-1";
    /**
     * 字段格式,默认0
     */
    private String format;
    /**
     * 子字段字段，如果没有则默认使用
     */
    private List<ExcelSubParam> subParam = new ArrayList<ExcelSubParam>();

    public String getNameCn() {
        return nameCn;
    }

    public void setNameCn(String nameCn) {
        this.nameCn = nameCn;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public List<ExcelSubParam> getSubParam() {
        return subParam;
    }

    public void setSubParam(List<ExcelSubParam> subParam) {
        this.subParam = subParam;
    }
}
