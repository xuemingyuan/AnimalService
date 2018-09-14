package com.animal;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.List;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;

import com.animal.utils.StringUtil;
import com.animal.utils.XlsUtils;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Level;
import org.apache.log4j.Logger;

public class Tools {

	private static Logger logger = Logger.getRootLogger();
	static {
		BasicConfigurator.configure();
		Logger.getRootLogger().setLevel(Level.ALL);
	}

	public static String getProjectPath() throws UnsupportedEncodingException {

		java.net.URL url = Tools.class.getProtectionDomain().getCodeSource().getLocation();
		String filePath = java.net.URLDecoder.decode(url.getPath(), "utf-8");
		if (filePath.endsWith(".jar"))
			filePath = filePath.substring(0, filePath.lastIndexOf("/") + 1);
		java.io.File file = new java.io.File(filePath);
		filePath = file.getAbsolutePath();
		return filePath;
	}

	public static void main(String[] args) throws Exception {

		logger.debug(" Begin. please wait one minute.");
		// D:/Java/ Monkey
		logger.debug(getProjectPath());
		String mainPath = getProjectPath();
		String kind = args[0];
		 kind = "Monkey";
		String path = mainPath + "/xls/test.xls";
		List<String> xlsList = XlsUtils.readXls(path);
		Tools f = new Tools();
		f.testWrite(xlsList, mainPath, kind);

		logger.debug(" All are finished. ");
	}

	public void testWrite(List<String> xlsList, String mainPath, String kind) throws Exception {
		String tempDir = mainPath + "/template";

		clearMainPathDoc(mainPath + "/doc");

		File directory = new File(tempDir);
		for (File templateFile : directory.listFiles()) {

			String allType = "only";
			String templatePath = templateFile.getPath();
			if (templatePath.contains("&"))
				allType = "both";
			for (String row : xlsList) {
				try (InputStream is = new FileInputStream(templatePath); HWPFDocument doc = new HWPFDocument(is);) {

					Range range = doc.getRange();
					String[] arr = row.split(",");
					range.replaceText("${animalNo}", StringUtil.getFixedLenString(arr[0], 6, ' '));
					range.replaceText("${tattooNo}", arr[1]);
					range.replaceText("${Sex}", arr[2]);
					range.replaceText("${Species}", kind);
					File dirPa = new File(mainPath + "/doc/" + kind + "/" + allType);
					dirPa.mkdirs();

					File file = new File(dirPa.getPath() + "/" + arr[0] + "_" + templateFile.getName());
					OutputStream os = new FileOutputStream(file);
					doc.write(os);
				}
			}
		}
	}

	private void clearMainPathDoc(String path) {

		File file = new File(path);
		if (file.isDirectory()) {
			for (File f : file.listFiles()) {
				clearMainPathDoc(f.getPath());
			}
		} else
			file.delete();

	}
}