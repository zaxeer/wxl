package com.zax.wxl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

import javax.swing.JTextPane;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class WordTemplateParser {
	private String filePath;
	private JTextPane output;
	
	public WordTemplateParser(String filePath) {
		super();
		this.filePath = filePath;
	}
	
	public String getFilePath() {
		return filePath;
	}
	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}
	public JTextPane getOutput() {
		return output;
	}
	public void setOutput(JTextPane output) {
		this.output = output;
	}
	
	public void parseWordFile(List<List<String>> replacements) throws InvalidFormatException, IOException {
		XWPFDocument doc = new XWPFDocument(OPCPackage.open(getFilePath()));
		for (XWPFParagraph p : doc.getParagraphs()) {
		    List<XWPFRun> runs = p.getRuns();
		    if (runs != null) {
		        for (XWPFRun r : runs) {
		            String text = r.getText(0);
		            if (text != null && text.contains("needle")) {
		                text = text.replace("needle", "haystack");
		                r.setText(text, 0);
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
		              if (text.contains("needle")) {
		                text = text.replace("needle", "haystack");
		                r.setText(text);
		              }
		            }
		         }
		      }
		   }
		}
		String directory = new File(filePath).getParent();
		doc.write(new FileOutputStream(directory +"/output.docx"));
	}

}
