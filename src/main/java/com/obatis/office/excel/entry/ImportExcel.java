package com.obatis.office.excel.entry;

import com.obatis.convert.date.DateCommonConvert;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class ImportExcel {

	private static final String MERGE_FLAG = "flag";
	private static final String MERGE_IS_VALUE = "isValue";
	private static final String MERGE_VALUE = "value";

	/**
	 * data为Excel流数据，表示从第2行开始读取数据，类似数组下标为1
	 * @param is
	 * @return
	 * @throws Exception
	 */
	public static List<List<String>> readExcel(InputStream is) throws Exception {
		// 表示默认从第2行开始读取，不读取第一行的标题
		return readExcel(is, 1);
	}

	/**
	 * data为Excel 字节流数数组，startRow表示从第几行开始读取数据，第几行类似数组下标
	 *
	 * @param is
	 * @param startRow
	 * @return
	 * @throws Exception
	 */
	public static List<List<String>> readExcel(InputStream is, int startRow) throws Exception {

		List<List<String>> list = new ArrayList<>();
		Workbook book = WorkbookFactory.create(is);
		if (book == null) {
			return list;
		}

		Sheet sheet = book.getSheetAt(0);
		if (sheet == null) {
			return list;
		}
		int rowNum = sheet.getLastRowNum();
		int cellNum = sheet.getRow(0).getLastCellNum();
		Map<String, String> mergeMap = new HashMap<>();
		for (int i = startRow; i <= rowNum; ++i) {
			Row row = sheet.getRow(i);
			if (row != null) {
				List<String> rowList = new ArrayList<>();
				int cellFlag = cellNum;
				for (int k = 0; k < cellNum; ++k) {
					Cell cell = row.getCell(k);

					if (cell != null) {
						Map<String, Object> mergeResult = isMergedRegion(sheet, cell, mergeMap);
						String content;
						if ((boolean) mergeResult.get(MERGE_FLAG)) {
							// 说明是合并单元格
							if ((boolean) mergeResult.get(MERGE_IS_VALUE)) {
								content = (String) mergeResult.get(MERGE_VALUE);
								if (content != null && "".equals(content)) {
									if(cellFlag == cellNum) {
										cellFlag = -1;
									}
									rowList.add(content);
								}
							}
						} else {
							content = getCellValue(cell);
							if (content == null || content.isEmpty()) {
								rowList.add("");
							} else {
								if(cellFlag == cellNum) {
									cellFlag = -1;
								}
								rowList.add(content);
							}
						}
					}else{
                        rowList.add("");
                    }

				}
				if (cellFlag != cellNum) {
					list.add(rowList);
				}
			}
		}
		return list;
	}

	private static Map<String, Object> isMergedRegion(Sheet sheet, Cell cell, Map<String, String> mergeMap) {
		Map<String, Object> result = new HashMap<>();
		boolean flag = false;
		boolean isValue = false;
		int row = cell.getRowIndex();
		int column = cell.getColumnIndex();

		int sheetMergeCount = sheet.getNumMergedRegions();
		for (int i = 0; i < sheetMergeCount; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			if (row >= firstRow && row <= lastRow) {
				if (column >= firstColumn && column <= lastColumn) {
					flag = true;
					String cellKey = firstColumn + "," + lastColumn + "," + firstRow + "," + lastRow;
					if (!mergeMap.containsKey(cellKey)) {
						mergeMap.put(cellKey, cellKey);
						isValue = true;
						result.put(MERGE_VALUE, getCellValue(cell));
					}
					break;
				}
			}
		}

		result.put(MERGE_FLAG, flag);
		result.put(MERGE_IS_VALUE, isValue);
		return result;

	}

    private static String getCellValue(Cell cell) {

		if (cell == null) {
			return "";
		}

		if (cell.getCellType() == CellType.STRING) {
			return cell.getStringCellValue();
		} else if (cell.getCellType() == CellType.BOOLEAN) {
			return String.valueOf(cell.getBooleanCellValue());
		} else if (cell.getCellType() == CellType.FORMULA) {
			return cell.getCellFormula();
		} else if (cell.getCellType() == CellType.NUMERIC) {
			String content = null;
            // 处理日期格式、时间格式
			if (HSSFDateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf;
				if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
					sdf = new SimpleDateFormat("HH:mm");
//					sdf = DefaultDateConstant.SD_FORMAT_HOUR_MINUTE;
				} else {
					// 日期

					sdf = new SimpleDateFormat("yyyy-MM-dd");
//					sdf = DefaultDateConstant.SD_FORMAT_DATE;
				}
				Date date = cell.getDateCellValue();
				content = sdf.format(date);
			} else if (cell.getCellStyle().getDataFormat() == 58) {
				// 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
//				content = DateCommonConvert.formatDate(DateUtil.getJavaDate(cell.getNumericCellValue()));
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				double value = cell.getNumericCellValue();
				Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
				content = sdf.format(date);
			} else {
				double value = cell.getNumericCellValue();
				CellStyle style = cell.getCellStyle();
				DecimalFormat format = new DecimalFormat();
				String temp = style.getDataFormatString();
				// 单元格设置成常规
				if (temp.equals("General")) {
					format.applyPattern("#");
				}
				content = format.format(value);
			}
			return content;
		}
		return "";
	}
//
//
//    /**
//     * 导入文件方法，需自定义接收类并使用ImpFiled注解
//     * @param f 导入文件
//     * @param cls 泛型类
//     * @param errorMsg 错误消息
//     * @param <T>
//     * @return
//     */
//	public static <T> List<T> excelToList(File f, Class<T> cls, List<String> errorMsg) {
//		List<T> list = ImportExcel.excelHandle(f, ImportExcel.getImpFiled(cls), cls, errorMsg);
//		return list;
//	}
//
//    /**
//     * 数据处理返回
//     * @param file
//     * @param fields
//     * @param cls
//     * @param errorMsg
//     * @param <T>
//     * @return
//     */
//	private static <T> List<T> excelHandle(File file, List<String> fields, Class<T> cls, List<String> errorMsg) {
//		List<T> array = new ArrayList<>();
//		List<JSONObject> varray = new ArrayList<>();
//		//错误提示信息
//		try {
//			InputStream in = new FileInputStream(file);
//			List<List<String>> list = ImportExcel.readExcel(in,1);
//			in.close();
//			for(int i = 0, j=list.size(); i<j; i++) {
//				List<String> rowValue = list.get(i);
//				JSONObject obj = new JSONObject();
//				if(fields.size() != rowValue.size()){
//					throw new RuntimeException("列数据与文件数据不一致");
//				}
//				for(int x = 0, y = rowValue.size(); x<y; x++){
//					obj.put(fields.get(x), rowValue.get(x));
//				}
//				varray.add(obj);
//				array.add(JSONObject.toJavaObject(obj, cls));
//			}
//			validateRow(varray, array, errorMsg);
//		} catch (FileNotFoundException e) {
//			e.printStackTrace();
//		} catch (IOException e) {
//			e.printStackTrace();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//		return  array;
//	}
//
//
//	/**
//	 * 获取导入字段
//	 * @param c
//	 * @param <T>
//	 * @return
//	 */
//    private static <T>  List<String> getImpFiled(Class<T> c) {
//		List<String> list = null;
//		try {
//			Constructor<?> constructor = c.getDeclaredConstructor();
//			Object p = constructor.newInstance();
//			Field[] fields = p.getClass().getDeclaredFields();
//			list = new ArrayList<>();
//			for(int i=0;i<fields.length-1;i++){
//				for(int j=0; j<fields.length-1-i; j++){
//	 				if(((ImpFiled)(fields[j].getDeclaredAnnotations()[0])).index() > ((ImpFiled)(fields[j+1].getDeclaredAnnotations()[0])).index()){
// 						Field temp = fields[j];
//	 					fields[j] = fields[j+1];
// 						fields[j+1] = temp;
//		 			}
//				}
//	 		}
//			for (Field field : fields){
//				//获得所有的注解
//				for (Annotation anno : field.getDeclaredAnnotations()) {
//					//找到自己的注解
//					if (anno.annotationType().equals(ImpFiled.class)) {
//						list.add(((ImpFiled)anno).index(), field.getName());
//					}
//				}
//			}
//		} catch (NoSuchMethodException e) {
//			e.printStackTrace();
//		} catch (IllegalAccessException e) {
//			e.printStackTrace();
//		} catch (InvocationTargetException e) {
//			e.printStackTrace();
//		} catch (InstantiationException e) {
//			e.printStackTrace();
//		}
//
//		return list;
//	}

//
//    /**
//     * 验证数据
//     * @param varray
//     * @param list
//     * @param errorMsg
//     * @param <T>
//     */
//    private static <T> void validateRow(List<JSONObject> varray, List<T> list, List<String> errorMsg){
//		for(int i = 0 ; i < varray.size(); i++ ){
//			try {
//				validateAnno(varray.get(i), list.get(i));
//			}catch (RuntimeException e){
//				e.printStackTrace();
//				errorMsg.add("第("+ (i+1) +")行导入数据有误，"+e.getMessage());
//			}catch (Exception e){
//				e.printStackTrace();
//				errorMsg.add("请检查第【"+i+1+"】行数据的数值是否正确！");
//			}
//		}
//	}
//
//    /**
//     * 验证注解
//     * @param jsonObject
//     * @param object
//     * @param <T>
//     */
//    private static <T> void validateAnno(JSONObject jsonObject, T object){
//		Set<String> keys = jsonObject.keySet();
//		for(String key : keys){
//			try {
//				Field f = object.getClass().getDeclaredField(key);
//				Annotation[] annos = f.getAnnotations();
//				Object value = jsonObject.get(key).toString().trim().equals("")  ? null : jsonObject.get(key).toString().trim();
//				for(Annotation annotation :annos){
//					if(annotation instanceof ImpFiled){
//						ImpFiled anno = (ImpFiled)annotation;
//						if(anno != null && anno.value().length() > 0){
//							JSONObject jsonValue = JSONObject.parseObject(anno.value());
//                            if(anno.notNull() && (null == value)){
//                                throw new RuntimeException("[" + anno.fieldName() + "]不能为空");
//                            }
//							value = jsonValue.get(value);
//							if(null == value){
//								throw new RuntimeException("[" + anno.fieldName() + "]格式应为：" + anno.value() + "，请填写对应参数");
//							}
//						}
//					}
//				}
//			} catch (NoSuchFieldException e) {
//				e.printStackTrace();
//			}
//
//		}
//	}

}