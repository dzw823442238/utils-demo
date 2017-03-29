package com.dary.poi.excl;

import java.io.File;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.junit.Test;

import com.dary.utils.poi.excl.ExclImportUtil;

/**  
* @Title: importUtilTest.java
* @Package com.dary.poi.excl
* @Description: ImportUtil.java测试类
* @author zhiwei.deng
* @date 2017年3月29日 下午3:35:43
* @version V1.0  
*/
public class importUtilTest {

	public static File file = new File("C:\\Users\\hp\\Desktop\\test.xlsx");
	
	@Test
	public void testGetCellValue() throws Exception{
		Sheet sheet = ExclImportUtil.getExclSheet(file);
		List<Row> list = ExclImportUtil.getExclSheetRows(sheet, 1);
		list.forEach(row ->{
			System.out.print(ExclImportUtil.getCellValue(row, 0)+"\t");
			System.out.print(ExclImportUtil.getCellValue(row, 1)+"\t");
			System.out.println(ExclImportUtil.getCellValue(row, 2)+"\t");
		});
	}
	
	@Test
	public void testImportExcl() throws Exception{
		String[][] strings = ExclImportUtil.importFile(file);
		for (String[] strings2 : strings) {
			for (String string : strings2) {
				System.out.print(string+"\t\t");
			}
			System.out.println();
		}
	}
}
