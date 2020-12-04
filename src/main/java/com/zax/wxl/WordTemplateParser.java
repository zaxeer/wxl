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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import javax.swing.JTextPane;
import javax.swing.SwingUtilities;
import javax.swing.text.BadLocationException;
import javax.swing.text.SimpleAttributeSet;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordTemplateParser {
	private JTextPane output;
	private String optionalWordTemplate;
	private boolean useOptionalWordTemplate = false;

	public WordTemplateParser() {
		super();
	}

	public JTextPane getOutput() {
		return output;
	}

	public void setOutput(JTextPane output) {
		this.output = output;
	}

	public void parseWordFile(List<List<String>> replacements, List<Integer> columns)
			throws InvalidFormatException, IOException, BadLocationException {

		int templatePathIndex = replacements.get(0).indexOf("WORD_TEMPLATE");

		if (StringUtils.isNotBlank(optionalWordTemplate)
				&& (optionalWordTemplate.endsWith("doc") || optionalWordTemplate.endsWith("docx"))
				&& Files.exists(Paths.get(optionalWordTemplate))) {
			useOptionalWordTemplate = true;
		}

		for (int count = 1; count < replacements.size(); count++) {
			String templatePath = replacements.get(count).get(templatePathIndex);
			if (useOptionalWordTemplate) {
				templatePath = optionalWordTemplate;
			}

			if (StringUtils.isBlank(templatePath) || !Files.exists(Paths.get(templatePath))
					|| (!templatePath.endsWith("doc") && !templatePath.endsWith("docx"))) {
				throw new FileNotFoundException("Word template not found or not valid " + templatePath);
			}

			String fileName = "";
			boolean addUnderscore = false;
			for (Integer col : columns) {
				int colValue = col.intValue();
				colValue--;
				if (addUnderscore) {
					fileName += "_" + replacements.get(count).get(colValue);
				} else {
					fileName += replacements.get(count).get(colValue);
				}
				addUnderscore = true;
			}

			try {
				XWPFDocument doc = new XWPFDocument(OPCPackage.open(templatePath));// don't close will update template
				for (XWPFParagraph p : doc.getParagraphs()) {
					List<XWPFRun> runs = p.getRuns();
					if (runs != null) {
						for (XWPFRun r : runs) {
							String text = r.getText(0);
							if (text != null && !StringUtils.isBlank(text)) {
								r.setText(replaceSearch(text, replacements.get(0), replacements.get(count)), 0);
							}
						}
					}
				}

				for (XWPFTable tbl : doc.getTables()) {
					for (XWPFTableRow row : tbl.getRows()) {
						for (XWPFTableCell cell : row.getTableCells()) {
							for (XWPFParagraph p : cell.getParagraphs()) {
								for (XWPFRun r : p.getRuns()) {
									String text = r.getText(0);
									if (text != null && !StringUtils.isBlank(text)) {
										r.setText(replaceSearch(text, replacements.get(0), replacements.get(count)));
									}
								}
							}
						}
					}
				}
				String directory = new File(templatePath).getParent();
				final String pathOut = directory + File.separator + count + "_"
						+ fileName.replaceAll("[\\\\/:\"*?<>|\\s]+", "_") + ".docx";

				doc.write(new FileOutputStream(pathOut));
				updateGUI(pathOut);

			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}

	private void updateGUI(final String pathOut) {
		SwingUtilities.invokeLater(new Runnable() {
			public void run() {
				try {
					output.getStyledDocument().insertString(output.getText().length(),
							"\nFile created -> " + pathOut, new SimpleAttributeSet());
				} catch (BadLocationException e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * @param text
	 * @param searchList
	 * @param replaceList
	 * @return
	 */
	private String replaceSearch(String text, List<String> searchList, List<String> replaceList) {
		for (int count = 0; count < searchList.size(); count++) {
			String search = searchList.get(count);
			String replace = replaceList.get(count);
			if (text.contains(search)) {
				text = text.replace(search, replace);
				System.out.println("founded in " + text);
				System.out.println("replacing " + search + " -> " + replace);
			}
		}
		return text;
	}

	public String getOptionalWordTemplate() {
		return optionalWordTemplate;
	}

	public void setOptionalWordTemplate(String optionalWordTemplate) {
		this.optionalWordTemplate = optionalWordTemplate;
	}

}
