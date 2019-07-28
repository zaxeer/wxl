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
import javax.swing.SwingWorker;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

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
	private JButton excelButton;
	private JButton btnStart;
	private JTextPane txtpnOutputAppearHere = new JTextPane();
	private JTextField textFieldFileName;

	/**
	 * Create the frame.
	 */
	public MainGUI() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 796, 376);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);

		excelButton = new JButton("Browse");
		excelButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				File workingDirectory = new File(System.getProperty("user.dir"));
				jFileChooser.setCurrentDirectory(workingDirectory);
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx", "xls");
				jFileChooser.setFileFilter(filter);
				jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
				int returnVal = jFileChooser.showOpenDialog((Component) e.getSource());
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					excelListPath.setText(jFileChooser.getSelectedFile().getPath());
				}
			}
		});

		excelListPath = new JTextField();
		excelListPath.setText("I:\\Jamati Work\\Taleem UL Quraan\\Letters for Manzoori\\List.xlsx");
		excelListPath.setColumns(10);

		lblPleaseSelectExcel = new JLabel("Please Select Excel List:");

		JScrollPane scrollPane = new JScrollPane();

		btnStart = new JButton("Start");
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				SwingWorker<String, Void> myWorker = new SwingWorker<String, Void>() {
					@Override
					protected String doInBackground() throws Exception {
						ExcelListReader excelListReader = new ExcelListReader(excelListPath.getText());
						WordTemplateParser templateParser = new WordTemplateParser();
						excelListReader.setOutput(txtpnOutputAppearHere);
						templateParser.setOutput(txtpnOutputAppearHere);
						String columsNumber = textFieldFileName.getText();
						String[] splited = columsNumber.split(",");
						List<Integer> columns = new ArrayList<Integer>();
						for(int count=0; count < splited.length ;count++) {
							columns.add(Integer.parseInt(splited[count].trim()));
						}
						try {
							templateParser.parseWordFile(excelListReader.parseExcel(),columns);
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
				"Create Last column in Excel file named as \"WORD_TEMPLATE\" to select template for each file.");
		lblCreateAColumn.setFont(new Font("Tahoma", Font.BOLD, 12));
		
		JLabel lblColumnsAsFile = new JLabel("Columns as File Name:");
		
		textFieldFileName = new JTextField();
		textFieldFileName.setText("1,2");
		textFieldFileName.setColumns(10);
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.TRAILING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addContainerGap()
							.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 754, Short.MAX_VALUE))
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblOutput))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(4)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(lblCreateAColumn)
										.addGroup(gl_contentPane.createSequentialGroup()
											.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
												.addComponent(lblPleaseSelectExcel)
												.addGroup(gl_contentPane.createSequentialGroup()
													.addGap(10)
													.addComponent(lblColumnsAsFile)))
											.addGap(36)
											.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
												.addComponent(textFieldFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
												.addComponent(excelListPath, GroupLayout.DEFAULT_SIZE, 534, Short.MAX_VALUE))))))
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING, false)
								.addComponent(excelButton, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(btnStart, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
					.addGap(6))
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addComponent(lblCreateAColumn)
					.addGap(15)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblPleaseSelectExcel)
						.addComponent(excelListPath, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(excelButton))
					.addPreferredGap(ComponentPlacement.RELATED)
					.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING, false)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addComponent(btnStart)
							.addGap(11))
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
								.addComponent(lblColumnsAsFile)
								.addComponent(textFieldFileName, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
							.addPreferredGap(ComponentPlacement.RELATED, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
							.addComponent(lblOutput)
							.addPreferredGap(ComponentPlacement.RELATED)))
					.addGap(6)
					.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 212, Short.MAX_VALUE)
					.addContainerGap())
		);

		txtpnOutputAppearHere.setText("Output appear here.....");
		scrollPane.setViewportView(txtpnOutputAppearHere);
		contentPane.setLayout(gl_contentPane);
	}
}
