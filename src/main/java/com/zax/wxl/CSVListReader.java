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

import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.text.BadLocationException;
import javax.swing.text.SimpleAttributeSet;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvValidationException;

public class CSVListReader {

	private String filePath;
	private JTextPane output;

	public CSVListReader(String filePath) {
		super();
		this.filePath = filePath;
	}

	public String getFilePath() {
		return filePath;
	}

	public JTextPane getOutput() {
		return output;
	}

	public List<List<String>> parseCSV() throws IOException, CsvValidationException {

		List<List<String>> records = new ArrayList<List<String>>();
		try (BufferedReader br = new BufferedReader(new FileReader(this.getFilePath(), StandardCharsets.UTF_8))) {
			String line;
			while ((line = br.readLine()) != null) {
				// Auto-detect delimiter
				char delimiter = line.contains(";") ? ';' : ',';
				CSVParser parser = new CSVParserBuilder().withSeparator(delimiter).build();
				CSVReader reader = new CSVReaderBuilder(new StringReader(line)).withCSVParser(parser).build();

				String[] values = reader.readNext();
				if (values != null) {
					records.add(Arrays.asList(values));
					addToGUI(values);
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}

		addToGUILast();
		return records;
	}

	private void addToGUI(final String[] values) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					getOutput().getStyledDocument().insertString(output.getText().length(), Arrays.toString(values),
							new SimpleAttributeSet());
				} catch (BadLocationException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		});
	}

	private void addToGUILast() {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					getOutput().getStyledDocument().insertString(output.getText().length(),
							"\n\nExcel parsed... Preparing file creations . . .\n", new SimpleAttributeSet());
				} catch (BadLocationException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		});
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

	public void setOutput(JTextPane output) {
		this.output = output;
	}

}
