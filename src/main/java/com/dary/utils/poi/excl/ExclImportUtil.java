package com.dary.utils.poi.excl;

import java.io.File;
import java.nio.file.Files;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**  
 * 解析excl文件，
* @Title: ImportUtil.java
* @Package com.dary.utils.poi.excl
* @Description: <p>Excl文件导入工具类  <br>支持2003,2007Excl格式</p>
* @author zhiwei.deng
* @date 2017年3月29日 上午10:13:53
* @version V1.0  
*/ 
public class ExclImportUtil {
	
	/**
	 * 默认去掉第一行数据，
	 * @param file excl文件
	 * @return string[row][cell]
	 * @throws Exception
	 */
	public static String[][] importFile(File file) throws Exception{
		return importFile(file, 0, 1, null);
	}
	
	/**
	 * excel文件导入
	 * @param file 
	 * @param sheetIndex 第几个sheet
	 * @param ignoreRows 需要忽略掉的行数
	 * @return
	 * @throws Exception 
	 */
	public static String[][] importFile(File file,int sheetIndex, int ignoreRows,Integer colCnt) throws Exception {
		List<String[]> result = new ArrayList<String[]>();
		Workbook wb = getExclWorkbook(file);
		int rowSize = 0;

		Sheet st = wb.getSheetAt(sheetIndex);
		for (int rowIndex = ignoreRows; rowIndex <= st.getLastRowNum(); rowIndex++) {
			Row row = st.getRow(rowIndex);
			if (null == row) {
				continue;
			}
//			int tempRowSize = row.getLastCellNum() + 1;
			int tempRowSize = row.getPhysicalNumberOfCells();//cell个数
			if(colCnt==null){
				if (tempRowSize > rowSize) {
					rowSize = tempRowSize;
				}
				colCnt=rowSize;
			}
			String[] values = new String[colCnt];
			Arrays.fill(values, "");
			boolean hasValue = false; //false:不保存
			B:for (short columnIndex = 0; columnIndex < colCnt; columnIndex++) {
				String value = "";
				value = getCellValue(row, columnIndex);
				// 第一个cell，数据为空，跳过这行
				if (columnIndex == 0 && value.trim().equals("")) {
					hasValue = false;
					break B;
				}
				values[columnIndex] = StringUtils.trim(value);
				hasValue = true;
			}
			if (hasValue) {
				result.add(values);
			}
		}
			
		String[][] returnArray = new String[result.size()][colCnt];
		for (int i = 0; i < returnArray.length; i++) {
			returnArray[i] = (String[]) result.get(i);
		}
		return returnArray;
	}

	/**
	 * 获取excl文件的workbook
	 * @param exclFile
	 * @return
	 * @throws Exception
	 */
	public static Workbook getExclWorkbook(File exclFile) throws Exception{
		Workbook workbook = null;
		if(!exclFile.exists())
			throw new NoSuchFieldException("指定文件不存在，请检查文件路径和文件名");
		try{
			workbook =(Workbook) new HSSFWorkbook(Files.newInputStream(exclFile.toPath()));	//2003
		}catch(Exception e){
			try{
				workbook =(Workbook) new XSSFWorkbook(Files.newInputStream(exclFile.toPath()));	//2007
			}catch(Exception e2){
				throw new Exception("未知的excl版本，请使用2003,2007excl版本",e);
			}
		}
		return workbook;
	}
	
	/**
	 * 得到workbook的sheet页
	 * @param workbook
	 * @param sheetIndex 下标
	 * @return
	 */
	public static Sheet getExclSheet(Workbook workbook,int sheetIndex){
		return workbook.getSheetAt(sheetIndex);
	}
	
	/**
	 * 获取文件的第一个sheet
	 * @param exclFile
	 * @return
	 * @throws Exception
	 */
	public static Sheet getExclSheet(File exclFile) throws Exception{
		return getExclSheet(getExclWorkbook(exclFile),0);
	}
	
	/**
	 * 获取文件的指定位置的sheet
	 * @param exclFile
	 * @return
	 * @throws Exception
	 */
	public static Sheet getExclSheet(File exclFile,int sheetIndex) throws Exception{
		return getExclSheet(getExclWorkbook(exclFile),sheetIndex);
	}
	/**
	 * 获取sheet中所有行rows
	 * @param sheet
	 * @param ignoreRows 需要忽略的行数 ，例如 5：表示忽略前5行,0:不忽略
	 * @return
	 */
	public static List<Row> getExclSheetRows(Sheet sheet,int ignoreRows){
		List<Row> list = new ArrayList<>();
		for (int rowIndex = ignoreRows; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
			Row row = sheet.getRow(rowIndex);
			if (null == row) {
				continue;
			}
			list.add(row);
		}
		return list;
	}
	
	/**
	 * 获取row中指定cell的值<br>
	 * excl cell 格式<br>
	 * 布尔值:输出Y/N <br>
	 * 日期：yyyy-MM-dd <br>
	 * @param row 
	 * @param cellIndex 下标位置
	 * @return
	 */
	public static String getCellValue(Row row,int cellIndex){
		String value = "";
		Cell cell = row.getCell(cellIndex);
		if (null != cell) {
			//根据cell类型格式化数据
			switch (cell.getCellType()) {
			case HSSFCell.CELL_TYPE_STRING:
				value = cell.getStringCellValue();
				break;
			case HSSFCell.CELL_TYPE_NUMERIC:
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					Date date = cell.getDateCellValue();
					if (null != date) {
						value = new SimpleDateFormat("yyyy-MM-dd").format(date);
					} else {
						value = "";
					}
				} else {
					DecimalFormat df = new DecimalFormat("0.#######");  								  
					//value = df.format(cell.getNumericCellValue()); 								
					value = String.valueOf(df.format(cell.getNumericCellValue()));
				}
				break;
			case HSSFCell.CELL_TYPE_FORMULA:
				// 公式生成的数据
				if (!"".equals(cell.getStringCellValue())) {
					value = cell.getStringCellValue();
				} else {
					value = cell.getNumericCellValue() + "";
				}
				break;
			case HSSFCell.CELL_TYPE_BLANK:
				break;
			case HSSFCell.CELL_TYPE_ERROR:
				value = "";
				break;
			case HSSFCell.CELL_TYPE_BOOLEAN:
				value = (cell.getBooleanCellValue() == true ? "Y" : "N");
				break;
			default:
				value = "";
			}
		}
		return value;
	}
	
	
	
}
