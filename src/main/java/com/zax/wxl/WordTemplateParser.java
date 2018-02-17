package com.zax.wxl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.JTextPane;
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

	public WordTemplateParser() {
		super();
	}

	public JTextPane getOutput() {
		return output;
	}

	public void setOutput(JTextPane output) {
		this.output = output;
	}

	public void parseWordFile(List<List<String>> replacements)
			throws InvalidFormatException, IOException, BadLocationException {
		
		for (int count = 1; count < replacements.size(); count++) {
			String templatePath = replacements.get(count).get(replacements.get(count).size()-1);
			@SuppressWarnings("resource")
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
			doc.write(new FileOutputStream(directory + "/" + count + "_output.docx"));
			this.output.getStyledDocument().insertString(this.output.getText().length(),
					"\nFile created -> " + directory + "/" + count + "_output.docx", new SimpleAttributeSet());
		}

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

}
