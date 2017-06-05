package com.zax.wxl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JTextPane;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelListReader {

	private String filePath;
	private JTextPane output;

	public ExcelListReader(String filePath) {
		super();
		this.filePath = filePath;
	}

	public String getFilePath() {
		return filePath;
	}

	public JTextPane getOutput() {
		return output;
	}

	public List<List<String>> parseExcel() throws IOException {
		List<List<String>> result = new ArrayList<List<String>>();
		FileInputStream excelFile = new FileInputStream(new File(this.getFilePath()));
		Workbook workbook = new XSSFWorkbook(excelFile);
		Sheet datatypeSheet = workbook.getSheetAt(0);
		Iterator<Row> iterator = datatypeSheet.iterator();

		List row = null;
		while (iterator.hasNext()) {

			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			row = new ArrayList<String>();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				// getCellTypeEnum shown as deprecated for version 3.15
				// getCellTypeEnum ill be renamed to getCellType starting from
				// version 4.0
				if (currentCell.getStringCellValue() != null) {
					row.add(currentCell.getStringCellValue());
				}

			}
			result.add(row);
		}
		getOutput().setText("Excel parsed\n");
		getOutput().setText(result.toString());
		return result;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public void setOutput(JTextPane output) {
		this.output = output;
	}

}
