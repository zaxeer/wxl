/*******************************************************************************
 * Copyright 2019 zaheer
 * 
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License.  You may obtain a copy
 * of the License at
 * 
 *   http://www.apache.org/licenses/LICENSE-2.0
 * 
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.  See the
 * License for the specific language governing permissions and limitations under
 * the License.
 ******************************************************************************/
package com.zax.wxl;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.text.BadLocationException;
import javax.swing.text.SimpleAttributeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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

		while (iterator.hasNext()) {

			Row currentRow = iterator.next();
			Iterator<Cell> cellIterator = currentRow.iterator();
			final List<String> row = new ArrayList<String>();
			while (cellIterator.hasNext()) {
				Cell currentCell = cellIterator.next();
				if (currentCell.getCellType().equals(CellType.STRING)) {
					row.add(currentCell.getStringCellValue().trim());
				} else if (currentCell.getCellType().equals(CellType.NUMERIC)) { 
					Double dVal = currentCell.getNumericCellValue();
					int val = dVal.intValue();
					row.add(""+val);
				} else if(currentCell.getCellType().equals(CellType.FORMULA)) { 
					Double dVal = currentCell.getNumericCellValue();
					int val = dVal.intValue();
					row.add(""+val);					
				}
				

			}
			result.add(row);
			SwingUtilities.invokeLater(new Runnable() {
				public void run() {
					try {
						getOutput().getStyledDocument().insertString(output.getText().length(),row.toString(),new SimpleAttributeSet());
					} catch (BadLocationException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			});

		}
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					getOutput().getStyledDocument().insertString(output.getText().length(),"\n\nExcel parsed... Preparing file creations . . .\n",new SimpleAttributeSet());
				} catch (BadLocationException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		});
		workbook.close();
		return result;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public void setOutput(JTextPane output) {
		this.output = output;
	}

}
