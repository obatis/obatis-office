package com.obatis.office.excel.param;

import com.obatis.office.excel.constant.ColumnTypeEnum;

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
    private ColumnTypeEnum type = ColumnTypeEnum.STRING;
    /**
     * 字段格式,默认0
     */
    private String format;
    /**
     * 子字段字段，如果没有则默认使用
     */
    private List<ExcelSubParam> subParam = new ArrayList<>();

    public String getNameCn() {
        return nameCn;
    }

    public void setNameCn(String nameCn) {
        this.nameCn = nameCn;
    }

    public ColumnTypeEnum getType() {
        return type;
    }

    public void setType(ColumnTypeEnum type) {
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
