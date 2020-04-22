package com.obatis.excel.entry;

import com.obatis.excel.param.ExcelParam;
import com.obatis.excel.param.ExcelSubParam;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class Test {

    public static void main(String[] args) throws Exception {

        ExcelParam param = new ExcelParam();
        param.setHeader("测试导出Excel");

        List<ExcelSubParam> column = new ArrayList<>();
        ExcelSubParam nameSubParam = new ExcelSubParam();
        nameSubParam.setNameCn("姓名");
        column.add(nameSubParam);

        ExcelSubParam hbSubParam = new ExcelSubParam();
        hbSubParam.setNameCn("合并列");

        List<ExcelSubParam> hbColumnArray = new ArrayList<>();
        ExcelSubParam hbAddress = new ExcelSubParam();
        hbAddress.setNameCn("地址");
        hbColumnArray.add(hbAddress);

        ExcelSubParam hbSex = new ExcelSubParam();
        hbSex.setNameCn("性别");
        hbColumnArray.add(hbSex);

        hbSubParam.setSubParam(hbColumnArray);


        column.add(hbSubParam);

        ExcelSubParam beizhuSubParam = new ExcelSubParam();
        beizhuSubParam.setNameCn("备注");
        column.add(beizhuSubParam);

        param.setColumn(column);


        List<List<String>> list = new ArrayList<>();
        List<String> dataItem = new ArrayList<>();
        dataItem.add("小明");
        dataItem.add("贵阳");
        dataItem.add("男");
        dataItem.add("备注列");
        list.add(dataItem);


        HSSFWorkbook workbook = ExportExcel.exportExcel(param, list);

        //文档输出
        FileOutputStream out = new FileOutputStream("//Users/huanglongpu/Documents/work/" + new SimpleDateFormat("yyyyMMddHHmmss").format(new Date()).toString() +".xls");
        workbook.write(out);
        out.close();
    }
}
