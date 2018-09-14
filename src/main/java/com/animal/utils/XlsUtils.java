package com.animal.utils;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Pattern;

import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.animal.Tools;

/**
 * ����poi��ȡexcel�ļ�����
 * 
 * @author lihui
 *
 */
public class XlsUtils {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println(isNumber("111"));
	}

	public static void print1(String path) throws Exception {
		InputStream is = new FileInputStream(path);
		HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(is));
		// A text extractor for Excel files.
		// Returns the textual content of the file, suitable for indexing by
		// something like Lucene,
		// but not really intended for display to the user.
		// �����������excel�ļ������ݣ���ʾΪ�ַ���
		ExcelExtractor extractor = new ExcelExtractor(wb);
		// �ַ��������������ͣ����api
		extractor.setIncludeSheetNames(true);
		extractor.setFormulasNotResults(false);
		extractor.setIncludeCellComments(true);
		// ����ַ�����ʽ
		String text = extractor.getText();
		System.out.println(text);
	}

	public static List<String> readXls(String path) throws Exception {
		List<String> xlsList = new ArrayList<>();
		InputStream is = new FileInputStream(path);
		HSSFWorkbook wb = new HSSFWorkbook(is);
		HSSFSheet sheet = wb.getSheetAt(0);

		for (Iterator<Row> iter = (Iterator<Row>) sheet.rowIterator(); iter.hasNext();) {
			Row row = iter.next();
			Cell cell1 = row.getCell(0);
			Cell cell2 = row.getCell(1);
			Cell cell3 = row.getCell(2);
			if (null == cell1)
				continue;
			String str1 = toString(cell1);
			if( !isNumber(str1))
			{
				continue;
			}
			if (str1.contains("/"))
				continue;
			StringBuilder sb = new StringBuilder();
			sb.append(str1 + "," + toString(cell2) + "," + toString(cell3));
			xlsList.add(sb.toString());
		}
		is.close();
		return xlsList;
	}

	private static boolean isNumber(String str) {
		return Pattern.matches("^[0-9]++$", str);
	}

	private static String toString(Cell cell1) {
		return cell1.getCellType() == HSSFCell.CELL_TYPE_NUMERIC ? (cell1.getNumericCellValue() + "").replace(".0", "")
				: cell1.getStringCellValue();
	}

}