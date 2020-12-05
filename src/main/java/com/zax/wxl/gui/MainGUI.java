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
package com.zax.wxl.gui;

import java.awt.Component;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.SwingConstants;
import javax.swing.SwingWorker;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.zax.wxl.CSVListReader;
import com.zax.wxl.ExcelListReader;
import com.zax.wxl.WordTemplateParser;

public class MainGUI extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 4469094356311946377L;
	private JPanel contentPane;
	private JLabel lblPleaseSelectExcel;
	private JTextField excelListPath;
	private JButton btnWordTempate;
	private JButton excelButton;
	private JButton btnStart;
	private JTextPane txtpnOutputAppearHere = new JTextPane();
	private JTextField textFieldFileName;
	private JTextField textFieldWordTemplate;

	/**
	 * Create the frame.
	 */
	public MainGUI() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 1050, 486);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);

		excelButton = new JButton("Browse");
		excelButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				File workingDirectory = new File(System.getProperty("user.dir"));
				jFileChooser.setCurrentDirectory(workingDirectory);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel or CSV Files", "xlsx", "xls", "csv");
				jFileChooser.setFileFilter(filter);
				jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
				int returnVal = jFileChooser.showOpenDialog((Component) e.getSource());
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					excelListPath.setText(jFileChooser.getSelectedFile().getPath());
				}
			}
		});
		
		btnWordTempate = new JButton("Browse");
		btnWordTempate.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				File workingDirectory = new File(System.getProperty("user.dir"));
				jFileChooser.setCurrentDirectory(workingDirectory);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Word Files", "doc", "docx");
				jFileChooser.setFileFilter(filter);
				jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
				int returnVal = jFileChooser.showOpenDialog((Component) e.getSource());
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					textFieldWordTemplate.setText(jFileChooser.getSelectedFile().getPath());
				}
			}
		});

		excelListPath = new JTextField();
		excelListPath.setColumns(10);
		
		textFieldWordTemplate  = new JTextField();
		textFieldWordTemplate.setColumns(10);

		lblPleaseSelectExcel = new JLabel("Please Select Excel List:");
		lblPleaseSelectExcel.setHorizontalAlignment(SwingConstants.RIGHT);

		JScrollPane scrollPane = new JScrollPane();

		btnStart = new JButton("Start");
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				SwingWorker<String, Void> myWorker = new SwingWorker<String, Void>() {
					@Override
					protected String doInBackground() throws Exception {
						List<List<String>> data = null;
						if(excelListPath.getText().endsWith("csv")) {
							CSVListReader csvListReader = new CSVListReader(excelListPath.getText());
							csvListReader.setOutput(txtpnOutputAppearHere);
							data = csvListReader.parseCSV();
						} else {
							ExcelListReader excelListReader = new ExcelListReader(excelListPath.getText());
							excelListReader.setOutput(txtpnOutputAppearHere);
							data = excelListReader.parseExcel();
						}
						
						WordTemplateParser wordTemplateParser = new WordTemplateParser();						
						wordTemplateParser.setOutput(txtpnOutputAppearHere);
						wordTemplateParser.setOptionalWordTemplate(textFieldWordTemplate.getText());
						
						String columsNumber = textFieldFileName.getText();
						String[] splited = columsNumber.split(",");
						List<Integer> columns = new ArrayList<Integer>();
						for(int count=0; count < splited.length ;count++) {
							columns.add(Integer.parseInt(splited[count].trim()));
						}
						try {
							wordTemplateParser.parseWordFile(data,columns);
						} catch (IOException e) {
							txtpnOutputAppearHere.setText(e.getMessage());
							e.printStackTrace();
						} catch (InvalidFormatException e) {
							txtpnOutputAppearHere.setText(e.getMessage());
							e.printStackTrace();
						} catch (BadLocationException e) {
							txtpnOutputAppearHere.setText(e.getMessage());
							e.printStackTrace();
						}
						return null;
					}
				};
				myWorker.execute();

			}
		});

		JLabel lblOutput = new JLabel("Output:");

		JLabel lblCreateAColumn = new JLabel(
				"Create Last column in Excel file named as \"WORD_TEMPLATE\" to select template for each Row OR select one singel Word Template for all Rows");
		lblCreateAColumn.setVerticalAlignment(SwingConstants.TOP);
		lblCreateAColumn.setFont(new Font("Tahoma", Font.BOLD, 12));
		
		JLabel lblColumnsAsFile = new JLabel("Columns used to output File Names:");
		lblColumnsAsFile.setHorizontalAlignment(SwingConstants.RIGHT);
		
		textFieldFileName = new JTextField();
		textFieldFileName.setText("1,2");
		textFieldFileName.setColumns(10);
		
		JLabel lblWordTemplateLabel = new JLabel("Optional: Select one Word Template:");
		lblWordTemplateLabel.setHorizontalAlignment(SwingConstants.RIGHT);
		lblWordTemplateLabel.setVerticalAlignment(SwingConstants.TOP);		
		
		
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addContainerGap()
							.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 1008, Short.MAX_VALUE))
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(4)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(lblCreateAColumn)
										.addGroup(gl_contentPane.createSequentialGroup()
											.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING, false)
												.addComponent(lblWordTemplateLabel, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
												.addComponent(lblPleaseSelectExcel, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
												.addComponent(lblColumnsAsFile, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
											.addPreferredGap(ComponentPlacement.RELATED)
											.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
												.addComponent(textFieldFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
												.addComponent(excelListPath, Alignment.TRAILING, GroupLayout.DEFAULT_SIZE, 760, Short.MAX_VALUE)
												.addComponent(textFieldWordTemplate, GroupLayout.DEFAULT_SIZE, 760, Short.MAX_VALUE)))))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblOutput)))
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING)
								.addComponent(excelButton)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addComponent(btnStart, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
									.addPreferredGap(ComponentPlacement.RELATED))
								.addComponent(btnWordTempate))))
					.addContainerGap())
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addComponent(lblCreateAColumn, GroupLayout.PREFERRED_SIZE, 24, GroupLayout.PREFERRED_SIZE)
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblWordTemplateLabel)
						.addComponent(textFieldWordTemplate, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(btnWordTempate))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblPleaseSelectExcel)
						.addComponent(excelListPath, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(excelButton))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblColumnsAsFile)
						.addComponent(textFieldFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING)
						.addComponent(btnStart)
						.addComponent(lblOutput))
					.addGap(6)
					.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 284, Short.MAX_VALUE)
					.addContainerGap())
		);

		txtpnOutputAppearHere.setText("Output appear here.....");
		scrollPane.setViewportView(txtpnOutputAppearHere);
		contentPane.setLayout(gl_contentPane);
	}
}
