package com.obatis.excel.param;

import java.util.ArrayList;
import java.util.List;

/**
 * @author HuangLongPu
 * @date 2018年11月14日15:11:26
 */
public class ExcelParam {

    /**
     * 头标题
     */
    private String header;
    /**
     * 字符串信息描述显示
     */
    private String headerMidString;
    /**
     * 头部html片段显示
     */
    private String headerMidHtml;
    /**
     * 尾部html片段显示
     */
    private String headerEndHtml;
    /**
     * 列标题
     */
    private List<ExcelSubParam> column = new ArrayList<>();
    /**
     * 是否需要序号，false
     */
    private boolean isSerial = false;


    public String getHeader() {
        return header;
    }

    public void setHeader(String header) {
        this.header = header;
    }

    public List<ExcelSubParam> getColumn() {
        return column;
    }

    public void setColumn(List<ExcelSubParam> column) {
        this.column = column;
    }

    public boolean isSerial() {
        return isSerial;
    }

    public void setSerial(boolean serial) {
        isSerial = serial;
    }

    public String getHeaderMidString() {
        return headerMidString;
    }

    public void setHeaderMidString(String headerMidString) {
        this.headerMidString = headerMidString;
    }

    public String getHeaderMidHtml() {
        return headerMidHtml;
    }

    public void setHeaderMidHtml(String headerMidHtml) {
        this.headerMidHtml = headerMidHtml;
    }

    public String getHeaderEndHtml() {
        return headerEndHtml;
    }

    public void setHeaderEndHtml(String headerEndHtml) {
        this.headerEndHtml = headerEndHtml;
    }
}
