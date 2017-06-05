package com.zax.wxl.gui;

import java.awt.Component;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
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

import com.zax.wxl.ExcelListReader;

public class MainGUI extends JFrame {

	/**
	 * 
	 */
	private static final long serialVersionUID = 4469094356311946377L;
	private JPanel contentPane;
	private JTextField wordTemplatePath;
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
		setBounds(100, 100, 456, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		
		excelButton = new JButton("Browse");
		excelButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser jFileChooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel Files", "xlsx","xls");
				jFileChooser.setFileFilter(filter);
				jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
				int returnVal = jFileChooser.showOpenDialog((Component)e.getSource());
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					excelListPath.setText(jFileChooser.getSelectedFile().getPath());
				}
			}
		});
		
				JButton btnWordTemplate = new JButton("Browse");
				btnWordTemplate.addActionListener(new ActionListener() {
					public void actionPerformed(ActionEvent e) {
						JFileChooser jFileChooser = new JFileChooser();
						FileNameExtensionFilter filter = new FileNameExtensionFilter("Word Files", "docx","doc");
						jFileChooser.setFileFilter(filter);
						jFileChooser.removeChoosableFileFilter(jFileChooser.getAcceptAllFileFilter());
						int returnVal = jFileChooser.showOpenDialog((Component)e.getSource());
						if (returnVal == JFileChooser.APPROVE_OPTION) {
							wordTemplatePath.setText(jFileChooser.getSelectedFile().getPath());
						}
					}
				});
						
								JLabel lblPleaseSelectWord = new JLabel("Please Select Word Template:");
				
						wordTemplatePath = new JTextField();
						wordTemplatePath.setColumns(10);
				
				excelListPath = new JTextField();
				excelListPath.setColumns(10);
				
				lblPleaseSelectExcel = new JLabel("Please Select Excel List:");
		
		JScrollPane scrollPane = new JScrollPane();
		
		btnStart = new JButton("Start");
		btnStart.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				ExcelListReader excelListReader = new ExcelListReader(excelListPath.getText());
				excelListReader.setOutput(txtpnOutputAppearHere);
				try {
					excelListReader.parseExcel();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		});
		
		JLabel lblOutput = new JLabel("Output:");
		GroupLayout gl_contentPane = new GroupLayout(contentPane);
		gl_contentPane.setHorizontalGroup(
			gl_contentPane.createParallelGroup(Alignment.TRAILING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
						.addGroup(gl_contentPane.createSequentialGroup()
							.addContainerGap()
							.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 414, Short.MAX_VALUE))
						.addGroup(gl_contentPane.createSequentialGroup()
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
								.addGroup(gl_contentPane.createSequentialGroup()
									.addGap(4)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(lblPleaseSelectWord)
										.addComponent(lblPleaseSelectExcel))
									.addPreferredGap(ComponentPlacement.UNRELATED)
									.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING)
										.addComponent(excelListPath, GroupLayout.DEFAULT_SIZE, 194, Short.MAX_VALUE)
										.addComponent(wordTemplatePath, GroupLayout.DEFAULT_SIZE, 194, Short.MAX_VALUE)))
								.addGroup(gl_contentPane.createSequentialGroup()
									.addContainerGap()
									.addComponent(lblOutput)))
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(gl_contentPane.createParallelGroup(Alignment.LEADING, false)
								.addComponent(excelButton, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(btnWordTemplate, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
								.addComponent(btnStart, GroupLayout.DEFAULT_SIZE, GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))))
					.addGap(6))
		);
		gl_contentPane.setVerticalGroup(
			gl_contentPane.createParallelGroup(Alignment.LEADING)
				.addGroup(gl_contentPane.createSequentialGroup()
					.addContainerGap()
					.addGroup(gl_contentPane.createParallelGroup(Alignment.BASELINE)
						.addComponent(wordTemplatePath, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblPleaseSelectWord)
						.addComponent(btnWordTemplate))
					.addPreferredGap(ComponentPlacement.RELATED)
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
					.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 137, Short.MAX_VALUE)
					.addContainerGap())
		);
		
		
		txtpnOutputAppearHere.setText("Output appear here.....");
		scrollPane.setViewportView(txtpnOutputAppearHere);
		contentPane.setLayout(gl_contentPane);
	}
}
