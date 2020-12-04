/**
 * 
 */
package com.zax.wxl;

import java.awt.EventQueue;

import javax.swing.UIManager;

import com.zax.wxl.gui.MainGUI;

/**
 * @author zaheer
 *
 */
public class MainApp {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println("Starting........");		
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
					MainGUI frame = new MainGUI();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});	

	}

}
