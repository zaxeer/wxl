package com.zax.wxl.gui;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

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
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import com.zax.wxl.ExcelListReader;
import com.zax.wxl.WordTemplateParser;
import java.awt.Font;

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
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx","xls");
				jFileChooser.setFileFilter(filter);
				jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
				int returnVal = jFileChooser.showOpenDialog((Component)e.getSource());
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					excelListPath.setText(jFileChooser.getSelectedFile().getPath());
				}
			}
		});
				
				excelListPath = new JTextField();
				excelListPath.setText("D:\\Jamati Work\\Taleem UL Quraan\\Letter WQAF\\List.xlsx");
				excelListPath.setColumns(10);
				
				lblPleaseSelectExcel = new JLabel("Please Select Excel List:");
		
		JScrollPane scrollPane = new JScrollPane();
		
		btnStart = new JButton("Start");
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				ExcelListReader excelListReader = new ExcelListReader(excelListPath.getText());
				WordTemplateParser templateParser = new WordTemplateParser();
				excelListReader.setOutput(txtpnOutputAppearHere);
				templateParser.setOutput(txtpnOutputAppearHere);
				try {
					templateParser.parseWordFile(excelListReader.parseExcel());
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
			}
		});
		
		JLabel lblOutput = new JLabel("Output:");
		
		JLabel lblCreateAColumn = new JLabel("Create Last column in Excel file named as \"WORD_TEMPLATE\" to select template for each file.");
		lblCreateAColumn.setFont(new Font("Tahoma", Font.BOLD, 12));
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
											.addComponent(lblPleaseSelectExcel)
											.addGap(39)
											.addComponent(excelListPath, GroupLayout.DEFAULT_SIZE, 534, Short.MAX_VALUE)))))
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
					.addGroup(gl_contentPane.createParallelGroup(Alignment.TRAILING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addComponent(btnStart)
							.addGap(11))
						.addGroup(gl_contentPane.createSequentialGroup()
							.addComponent(lblOutput)
							.addPreferredGap(ComponentPlacement.RELATED)))
					.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 213, Short.MAX_VALUE)
					.addContainerGap())
		);
		
		
		txtpnOutputAppearHere.setText("Output appear here.....");
		scrollPane.setViewportView(txtpnOutputAppearHere);
		contentPane.setLayout(gl_contentPane);
	}
}
