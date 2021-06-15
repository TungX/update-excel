package vn.yinx.computeexcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.net.URLDecoder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

public class Main {
	public static void main(String[] args) throws Exception {
		JSONParser parser = new JSONParser();
		JSONObject input = (JSONObject) parser.parse(args[0]);
		String tempPath = input.get("template").toString();
		String outPath = input.get("output").toString();
		FileInputStream file = new FileInputStream(new File(tempPath));
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		JSONObject data = (JSONObject) input.get("data");
		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
		for(Object o: data.keySet()) {
			String key = o.toString();
			int c = key.charAt(0) - 65;
			int r = Integer.parseInt(key.substring(1)) - 1;
			Row row = sheet.getRow(r);
			if(row == null) {
				row = sheet.createRow(r);
			}
			Cell cell = row.getCell(c); 
			if(cell == null) {
				cell = row.createCell(c);
			}
			String value = data.get(o).toString();
			if(value.startsWith("url-") && value.endsWith("-encode")) {
				value = URLDecoder.decode(value.replace("url-", "").replace("-encode", ""), "UTF-8");
			}
			try {
				cell.setCellValue(Double.parseDouble(value));
			} catch (Exception e) {
				// TODO: handle exception
				cell.setCellValue(value);
			}
		}
//		for(int i = 0; i < 100; i++) {
//			Row row = sheet.getRow(i);
//			if(row == null){
//				continue;
//			}
//			Cell cell = row.getCell(0);
//			if(cell == null) {
//				continue;
//			}
//			if(cell.getCellType() == CellType.NUMERIC) {
//				System.out.println("Row at "+i+": "+cell.getNumericCellValue());
//				cell.setCellValue(11111.11);
//			}else if(cell.getCellType() == CellType.STRING) {
//				System.out.println("Row at "+i+": "+cell.getStringCellValue());
//			}
//		}
		file.close();
		FileOutputStream fileOut = new FileOutputStream(outPath);
		workbook.write(fileOut);
		fileOut.close();

		// Closing the workbook
		workbook.close();
		System.out.println("success");
	}
}
