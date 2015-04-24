import java.awt.CardLayout;
import java.awt.Component;
import java.awt.EventQueue;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class App {

	private JFrame frame;
	ArrayList<String> columns = new ArrayList<String>();
	FileOutputStream fos;
	private JTable table;
	Workbook workbook;
	Workbook workbook2;


	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					App window = new App();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public App() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	@SuppressWarnings("unchecked")
	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 572, 463);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(new CardLayout(0, 0));
		
		JPanel panel = new JPanel();
		frame.getContentPane().add(panel, "name_95082484594159");
		GridBagLayout gbl_panel = new GridBagLayout();
		gbl_panel.columnWidths = new int[]{0, 0};
		gbl_panel.rowHeights = new int[]{0, 300, 96, 0, 0};
		gbl_panel.columnWeights = new double[]{1.0, Double.MIN_VALUE};
		gbl_panel.rowWeights = new double[]{0.0, 1.0, 0.0, 0.0, Double.MIN_VALUE};
		panel.setLayout(gbl_panel);
		
		JLabel label = new JLabel("What column to find duplicate");
		GridBagConstraints gbc_label = new GridBagConstraints();
		gbc_label.insets = new Insets(0, 0, 5, 0);
		gbc_label.gridx = 0;
		gbc_label.gridy = 0;
		panel.add(label, gbc_label);
		
		JScrollPane scrollPane = new JScrollPane();
		GridBagConstraints gbc_scrollPane = new GridBagConstraints();
		gbc_scrollPane.fill = GridBagConstraints.BOTH;
		gbc_scrollPane.insets = new Insets(0, 0, 5, 0);
		gbc_scrollPane.gridx = 0;
		gbc_scrollPane.gridy = 1;
		panel.add(scrollPane, gbc_scrollPane);
		
		table = new JTable();
		scrollPane.setViewportView(table);

		JComboBox comboBox = new JComboBox();
		GridBagConstraints gbc_comboBox = new GridBagConstraints();
		gbc_comboBox.insets = new Insets(0, 0, 5, 0);
		gbc_comboBox.fill = GridBagConstraints.HORIZONTAL;
		gbc_comboBox.gridx = 0;
		gbc_comboBox.gridy = 2;
		panel.add(comboBox, gbc_comboBox);
		
		

		
		JPanel panel_1 = new JPanel();
		frame.getContentPane().add(panel_1, "name_96484279449973");
		
		
		JFileChooser fileChooser = new JFileChooser();
		panel_1.add(fileChooser);
		int response = fileChooser.showOpenDialog(panel_1);
		if(response == JFileChooser.APPROVE_OPTION){
			
		System.out.println("size "+columns.size());
		ArrayList<String> columns = new ArrayList<String>();
		File file = fileChooser.getSelectedFile();
		String newName = file.getName()+"Dup.xlsx";
		workbook = null;
		try {
			workbook = new XSSFWorkbook(new FileInputStream(file));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		workbook2 = new XSSFWorkbook();
		columns = processExcel.process(workbook,workbook2,file);

		for(int i=0;i<columns.size();i++){
			comboBox.addItem(new Integer(i));
			System.out.println(comboBox.getItemAt(i));
		}
		
		Object[][] numbers = new Object[1][columns.size()];
		String[] columnTable = new String[columns.size()];
		System.out.println("\n columns arraylist"+columns);
		for(int i=0; i<columns.size();i++)numbers[0][i]=i;
		for(int j=0; j<columns.size();j++){
			columnTable[j]=columns.get(j);
			System.out.println("columns each"+columnTable[j]);
		}
		table = new JTable(numbers,columnTable);
		scrollPane.setViewportView(table);
		resizeColumnWidth(table);
		
		
		JButton btnNewButton = new JButton("Generate Report");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
				try {
					fos = new FileOutputStream(newName);
					if(workbook != null && workbook2 != null && comboBox.getSelectedIndex() != -1){
						processExcel.findDup(workbook, workbook2, fos, comboBox.getSelectedIndex());
						JOptionPane.showMessageDialog(null, "Done finding duplicate", "Yay, java", JOptionPane.PLAIN_MESSAGE);
					}
				} catch (Exception e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
					JOptionPane.showMessageDialog(null, "Error", "Yay, java", JOptionPane.PLAIN_MESSAGE);

				}
			}
		});
		GridBagConstraints gbc_btnNewButton = new GridBagConstraints();
		gbc_btnNewButton.gridx = 0;
		gbc_btnNewButton.gridy = 3;
		panel.add(btnNewButton, gbc_btnNewButton);
		}else{
			System.out.println("error");
		}
		
	}
	public void resizeColumnWidth(JTable table) {
	    final TableColumnModel columnModel = table.getColumnModel();
	    for (int column = 0; column < table.getColumnCount(); column++) {
	        int width = 50; // Min width
	        for (int row = 0; row < table.getRowCount(); row++) {
	            TableCellRenderer renderer = table.getCellRenderer(row, column);
	            Component comp = table.prepareRenderer(renderer, row, column);
	            width = Math.max(comp.getPreferredSize().width, width);
	        }
	        columnModel.getColumn(column).setPreferredWidth(width);
	    }
	}
}
