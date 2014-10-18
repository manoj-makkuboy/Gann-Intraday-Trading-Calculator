package gann;
import java.awt.Button;
import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.border.EmptyBorder;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.awt.Window.Type;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.JToggleButton;
import javax.swing.JLabel;

public class GUI extends JFrame {

	
	public static double round(double value, int places) {
	    if (places < 0) throw new IllegalArgumentException();

	    BigDecimal bd = new BigDecimal(value);
	    bd = bd.setScale(places, BigDecimal.ROUND_HALF_UP);
	    return bd.doubleValue();
	}
	
	
	double value_0;
	 double  p[] = new double[60];
	 double resistance[] = new double[6];
	 double support[] = new double[6];
	 
	 
	 
	 
	private JPanel contentPane;
	private JTextField p34;
	private JTextField p31;
	private JTextField textField_2;
	private JTextField p37;
	private JTextField textField_4;
	private JTextField textField_5;
	private JTextField textField_6;
	private JTextField textField_7;
	private JTextField textField_8;
	private JTextField p13;
	private JTextField p15;
	private JTextField black_2;
	private JTextField black_3;
	private JTextField textField_14;
	private JTextField textField_15;
	private JTextField black_1;
	private JTextField p4;
	private JTextField p3;
	private JTextField p5;
	private JTextField black_4;
	private JTextField p28;
	private JTextField p40;
	private JTextField p11;
	private JTextField p1;
	private JTextField p2;
	private JTextField p6;
	private JTextField p19;
	private JTextField textField_28;
	private JTextField textField_29;
	private JTextField black_8;
	private JTextField p8;
	private JTextField p9;
	private JTextField p7;
	private JTextField black_5;
	private JTextField textField_35;
	private JTextField textField_36;
	private JTextField p25;
	private JTextField p23;
	private JTextField black_7;
	private JTextField black_6;
	private JTextField p21;
	private JTextField p49;
	private JTextField p43;
	private JTextField textField_44;
	private JTextField p46;
	private JTextField textField_46;
	private JTextField textField_47;
	private JTextField textField_48;
	private JTextField input_box;
	private JTextField support_1;
	private JTextField p17;
	private JTextField support_2;
	private JTextField support_3;
	private JTextField support_4;
	private JTextField support_5;
	private JTextField support_0;
	private JTextField resistance_0;
	private JTextField resistance_1;
	private JTextField resistance_2;
	private JTextField resistance_3;
	private JTextField resistance_4;
	private JTextField resistance_5;
	
	
	private JToggleButton tglbtnHighlight;
	/**
	 * Launch the application.
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		
		
		
		
		
		
		
		EventQueue.invokeLater(new Runnable() {
			
			
			
			
			public void run() {
				try {
					GUI frame = new GUI();
					frame.setVisible(true);
					frame.setResizable(false);
					
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public GUI() {
		setType(Type.UTILITY);
		setTitle("Gann Calculator");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 392, 679);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		p34 = new JTextField();
		p34.setEditable(false);
		p34.setHorizontalAlignment(SwingConstants.CENTER);
		p34.setBackground(Color.LIGHT_GRAY);
		p34.setBounds(166, 0, 55, 47);
		contentPane.add(p34);
		p34.setColumns(10);
		
		p31 = new JTextField();
		p31.setEditable(false);
		p31.setHorizontalAlignment(SwingConstants.CENTER);
		p31.setBackground(Color.LIGHT_GRAY);
		p31.setBounds(0, 0, 55, 47);
		contentPane.add(p31);
		p31.setColumns(10);
		
		
		
		textField_2 = new JTextField();
		textField_2.setEditable(false);
		textField_2.setHorizontalAlignment(SwingConstants.CENTER);
		textField_2.setColumns(10);
		textField_2.setBounds(56, 0, 55, 47);
		contentPane.add(textField_2);
		
		p37 = new JTextField();
		p37.setEditable(false);
		p37.setHorizontalAlignment(SwingConstants.CENTER);
		p37.setBackground(Color.LIGHT_GRAY);
		p37.setColumns(10);
		p37.setBounds(331, 0, 55, 47);
		contentPane.add(p37);
		
		textField_4 = new JTextField();
		textField_4.setEditable(false);
		textField_4.setHorizontalAlignment(SwingConstants.CENTER);
		textField_4.setColumns(10);
		textField_4.setBounds(276, 0, 55, 47);
		contentPane.add(textField_4);
		
		textField_5 = new JTextField();
		textField_5.setEditable(false);
		textField_5.setHorizontalAlignment(SwingConstants.CENTER);
		textField_5.setColumns(10);
		textField_5.setBounds(221, 0, 55, 47);
		contentPane.add(textField_5);
		
		textField_6 = new JTextField();
		textField_6.setEditable(false);
		textField_6.setHorizontalAlignment(SwingConstants.CENTER);
		textField_6.setColumns(10);
		textField_6.setBounds(110, 0, 55, 47);
		contentPane.add(textField_6);
		
		textField_7 = new JTextField();
		textField_7.setEditable(false);
		textField_7.setHorizontalAlignment(SwingConstants.CENTER);
		textField_7.setColumns(10);
		textField_7.setBounds(0, 46, 55, 47);
		contentPane.add(textField_7);
		
		textField_8 = new JTextField();
		textField_8.setEditable(false);
		textField_8.setHorizontalAlignment(SwingConstants.CENTER);
		textField_8.setColumns(10);
		textField_8.setBounds(331, 46, 55, 47);
		contentPane.add(textField_8);
		
		p13 = new JTextField();
		p13.setEditable(false);
		p13.setHorizontalAlignment(SwingConstants.CENTER);
		p13.setBackground(Color.LIGHT_GRAY);
		p13.setColumns(10);
		p13.setBounds(56, 46, 55, 47);
		contentPane.add(p13);
		
		p15 = new JTextField();
		p15.setEditable(false);
		p15.setHorizontalAlignment(SwingConstants.CENTER);
		p15.setColumns(10);
		p15.setBackground(Color.LIGHT_GRAY);
		p15.setBounds(166, 46, 55, 47);
		contentPane.add(p15);
		
		black_2 = new JTextField();
		black_2.setEditable(false);
		black_2.setHorizontalAlignment(SwingConstants.CENTER);
		black_2.setColumns(10);
		black_2.setBounds(110, 46, 55, 47);
		contentPane.add(black_2);
		
		black_3 = new JTextField();
		black_3.setEditable(false);
		black_3.setHorizontalAlignment(SwingConstants.CENTER);
		black_3.setColumns(10);
		black_3.setBounds(221, 46, 55, 47);
		contentPane.add(black_3);
		
	    p17 = new JTextField();
	    p17.setEditable(false);
	    p17.setHorizontalAlignment(SwingConstants.CENTER);
		p17.setBackground(Color.LIGHT_GRAY);
		p17.setColumns(50);
		p17.setBounds(276, 46, 55, 47);
		contentPane.add(p17);
		
		textField_14 = new JTextField();
		textField_14.setEditable(false);
		textField_14.setHorizontalAlignment(SwingConstants.CENTER);
		textField_14.setColumns(10);
		textField_14.setBounds(0, 92, 55, 47);
		contentPane.add(textField_14);
		
		textField_15 = new JTextField();
		textField_15.setEditable(false);
		textField_15.setHorizontalAlignment(SwingConstants.CENTER);
		textField_15.setColumns(10);
		textField_15.setBounds(331, 92, 55, 47);
		contentPane.add(textField_15);
		
		black_1 = new JTextField();
		black_1.setEditable(false);
		black_1.setHorizontalAlignment(SwingConstants.CENTER);
		black_1.setColumns(10);
		black_1.setBounds(56, 92, 55, 47);
		contentPane.add(black_1);
		
		p4 = new JTextField();
		p4.setEditable(false);
		p4.setHorizontalAlignment(SwingConstants.CENTER);
		p4.setColumns(10);
		p4.setBackground(Color.LIGHT_GRAY);
		p4.setBounds(166, 92, 55, 47);
		contentPane.add(p4);
		
		p3 = new JTextField();
		p3.setEditable(false);
		p3.setHorizontalAlignment(SwingConstants.CENTER);
		p3.setBackground(Color.LIGHT_GRAY);
		p3.setColumns(10);
		p3.setBounds(110, 92, 55, 47);
		contentPane.add(p3);
		
		p5 = new JTextField();
		p5.setEditable(false);
		p5.setHorizontalAlignment(SwingConstants.CENTER);
		p5.setBackground(Color.LIGHT_GRAY);
		p5.setColumns(10);
		p5.setBounds(221, 92, 55, 47);
		contentPane.add(p5);
		
		black_4 = new JTextField();
		black_4.setEditable(false);
		black_4.setHorizontalAlignment(SwingConstants.CENTER);
		black_4.setColumns(10);
		black_4.setBounds(276, 92, 55, 47);
		contentPane.add(black_4);
		
		p28 = new JTextField();
		p28.setEditable(false);
		p28.setHorizontalAlignment(SwingConstants.CENTER);
		p28.setBackground(Color.LIGHT_GRAY);
		p28.setColumns(10);
		p28.setBounds(0, 140, 55, 47);
		contentPane.add(p28);
		
		p40 = new JTextField();
		p40.setEditable(false);
		p40.setHorizontalAlignment(SwingConstants.CENTER);
		p40.setBackground(Color.LIGHT_GRAY);
		p40.setColumns(10);
		p40.setBounds(331, 140, 55, 47);
		contentPane.add(p40);
		
		p11 = new JTextField();
		p11.setEditable(false);
		p11.setHorizontalAlignment(SwingConstants.CENTER);
		p11.setBackground(Color.LIGHT_GRAY);
		p11.setColumns(10);
		p11.setBounds(56, 140, 55, 47);
		contentPane.add(p11);
		
		p1 = new JTextField();
		p1.setEditable(false);
		p1.setForeground(Color.WHITE);
		p1.setHorizontalAlignment(SwingConstants.CENTER);
		p1.setColumns(10);
		p1.setBackground(Color.GRAY);
		p1.setBounds(166, 140, 55, 47);
		contentPane.add(p1);
		
		
		p2 = new JTextField();
		p2.setEditable(false);
		p2.setHorizontalAlignment(SwingConstants.CENTER);
		p2.setBackground(Color.LIGHT_GRAY);
		p2.setColumns(10);
		p2.setBounds(110, 140, 55, 47);
		contentPane.add(p2);
		
		p6 = new JTextField();
		p6.setEditable(false);
		p6.setHorizontalAlignment(SwingConstants.CENTER);
		p6.setBackground(Color.LIGHT_GRAY);
		p6.setColumns(10);
		p6.setBounds(221, 140, 55, 47);
		contentPane.add(p6);
		
		p19 = new JTextField();
		p19.setEditable(false);
		p19.setHorizontalAlignment(SwingConstants.CENTER);
		p19.setBackground(Color.LIGHT_GRAY);
		p19.setColumns(10);
		p19.setBounds(276, 140, 55, 47);
		contentPane.add(p19);
		
		textField_28 = new JTextField();
		textField_28.setEditable(false);
		textField_28.setHorizontalAlignment(SwingConstants.CENTER);
		textField_28.setColumns(10);
		textField_28.setBounds(0, 186, 55, 47);
		contentPane.add(textField_28);
		
		textField_29 = new JTextField();
		textField_29.setEditable(false);
		textField_29.setHorizontalAlignment(SwingConstants.CENTER);
		textField_29.setColumns(10);
		textField_29.setBounds(331, 186, 55, 47);
		contentPane.add(textField_29);
		
		black_8 = new JTextField();
		black_8.setEditable(false);
		black_8.setHorizontalAlignment(SwingConstants.CENTER);
		black_8.setColumns(10);
		black_8.setBounds(56, 186, 55, 47);
		contentPane.add(black_8);
		
		p8 = new JTextField();
		p8.setEditable(false);
		p8.setHorizontalAlignment(SwingConstants.CENTER);
		p8.setColumns(10);
		p8.setBackground(Color.LIGHT_GRAY);
		p8.setBounds(166, 186, 55, 47);
		contentPane.add(p8);
		
		p9 = new JTextField();
		p9.setEditable(false);
		p9.setHorizontalAlignment(SwingConstants.CENTER);
		p9.setBackground(Color.LIGHT_GRAY);
		p9.setColumns(10);
		p9.setBounds(110, 186, 55, 47);
		contentPane.add(p9);
		
		p7 = new JTextField();
		p7.setEditable(false);
		p7.setHorizontalAlignment(SwingConstants.CENTER);
		p7.setBackground(Color.LIGHT_GRAY);
		p7.setColumns(10);
		p7.setBounds(221, 186, 55, 47);
		contentPane.add(p7);
		
		black_5 = new JTextField();
		black_5.setEditable(false);
		black_5.setHorizontalAlignment(SwingConstants.CENTER);
		black_5.setColumns(10);
		black_5.setBounds(276, 186, 55, 47);
		contentPane.add(black_5);
		
		textField_35 = new JTextField();
		textField_35.setEditable(false);
		textField_35.setHorizontalAlignment(SwingConstants.CENTER);
		textField_35.setColumns(10);
		textField_35.setBounds(0, 234, 55, 47);
		contentPane.add(textField_35);
		
		textField_36 = new JTextField();
		textField_36.setEditable(false);
		textField_36.setHorizontalAlignment(SwingConstants.CENTER);
		textField_36.setColumns(10);
		textField_36.setBounds(331, 234, 55, 47);
		contentPane.add(textField_36);
		
		p25 = new JTextField();
		p25.setEditable(false);
		p25.setHorizontalAlignment(SwingConstants.CENTER);
		p25.setBackground(Color.LIGHT_GRAY);
		p25.setColumns(10);
		p25.setBounds(56, 234, 55, 47);
		contentPane.add(p25);
		
		p23 = new JTextField();
		p23.setEditable(false);
		p23.setHorizontalAlignment(SwingConstants.CENTER);
		p23.setColumns(10);
		p23.setBackground(Color.LIGHT_GRAY);
		p23.setBounds(166, 234, 55, 47);
		contentPane.add(p23);
		
		black_7 = new JTextField();
		black_7.setEditable(false);
		black_7.setHorizontalAlignment(SwingConstants.CENTER);
		black_7.setColumns(10);
		black_7.setBounds(110, 234, 55, 47);
		contentPane.add(black_7);
		
		black_6 = new JTextField();
		black_6.setEditable(false);
		black_6.setHorizontalAlignment(SwingConstants.CENTER);
		black_6.setColumns(10);
		black_6.setBounds(221, 234, 55, 47);
		contentPane.add(black_6);
		
		p21 = new JTextField();
		p21.setEditable(false);
		p21.setHorizontalAlignment(SwingConstants.CENTER);
		p21.setBackground(Color.LIGHT_GRAY);
		p21.setColumns(10);
		p21.setBounds(276, 234, 55, 47);
		contentPane.add(p21);
		
		p49 = new JTextField();
		p49.setEditable(false);
		p49.setHorizontalAlignment(SwingConstants.CENTER);
		p49.setBackground(Color.LIGHT_GRAY);
		p49.setColumns(10);
		p49.setBounds(0, 281, 55, 47);
		contentPane.add(p49);
		
		p43 = new JTextField();
		p43.setEditable(false);
		p43.setHorizontalAlignment(SwingConstants.CENTER);
		p43.setBackground(Color.LIGHT_GRAY);
		p43.setColumns(10);
		p43.setBounds(331, 281, 55, 47);
		contentPane.add(p43);
		
		textField_44 = new JTextField();
		textField_44.setEditable(false);
		textField_44.setHorizontalAlignment(SwingConstants.CENTER);
		textField_44.setColumns(10);
		textField_44.setBounds(56, 281, 55, 47);
		contentPane.add(textField_44);
		
		p46 = new JTextField();
		p46.setEditable(false);
		p46.setHorizontalAlignment(SwingConstants.CENTER);
		p46.setColumns(10);
		p46.setBackground(Color.LIGHT_GRAY);
		p46.setBounds(166, 281, 55, 47);
		contentPane.add(p46);
		
		textField_46 = new JTextField();
		textField_46.setEditable(false);
		textField_46.setHorizontalAlignment(SwingConstants.CENTER);
		textField_46.setColumns(10);
		textField_46.setBounds(110, 281, 55, 47);
		contentPane.add(textField_46);
		
		textField_47 = new JTextField();
		textField_47.setEditable(false);
		textField_47.setHorizontalAlignment(SwingConstants.CENTER);
		textField_47.setColumns(10);
		textField_47.setBounds(221, 281, 55, 47);
		contentPane.add(textField_47);
		
		textField_48 = new JTextField();
		textField_48.setEditable(false);
		textField_48.setHorizontalAlignment(SwingConstants.CENTER);
		textField_48.setColumns(10);
		textField_48.setBounds(276, 281, 55, 47);
		contentPane.add(textField_48);
		
		input_box = new JTextField();
		input_box.setBounds(10, 493, 60, 20);
		contentPane.add(input_box);
		input_box.setColumns(10);
		
		
		support_1 = new JTextField();
		support_1.setColumns(10);
		support_1.setBounds(99, 544, 55, 20);
		contentPane.add(support_1);
		
		Button btnCalculate = new Button("Gann Calculator");
		btnCalculate.setFont(UIManager.getFont("InternalFrame.titleFont"));
		btnCalculate.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
				
				
				
				black_1.setText(" ");black_1.setBackground(Color.WHITE);black_1.setForeground(Color.BLACK);
				
				black_2.setText(" ");black_2.setBackground(Color.WHITE);black_2.setForeground(Color.BLACK);
				black_3.setText(" ");black_3.setBackground(Color.WHITE);black_3.setForeground(Color.BLACK);
				black_4.setText(" ");black_4.setBackground(Color.WHITE);black_4.setForeground(Color.BLACK);
				black_5.setText(" ");black_5.setBackground(Color.WHITE);black_5.setForeground(Color.BLACK);
				black_6.setText(" ");black_6.setBackground(Color.WHITE);black_6.setForeground(Color.BLACK);
				black_7.setText(" ");black_7.setBackground(Color.WHITE);black_7.setForeground(Color.BLACK);
				black_8.setText(" ");black_8.setBackground(Color.WHITE);black_8.setForeground(Color.BLACK);
				
				
				p2.setForeground(Color.BLACK);
				p3.setForeground(Color.BLACK);
				p4.setForeground(Color.BLACK);
				p5.setForeground(Color.BLACK);
				p6.setForeground(Color.BLACK);
				p7.setForeground(Color.BLACK);
				p8.setForeground(Color.BLACK);
				p9.setForeground(Color.BLACK);
				
				p11.setForeground(Color.BLACK);
				p13.setForeground(Color.BLACK);
				p15.setForeground(Color.BLACK);
				p17.setForeground(Color.BLACK);
				p19.setForeground(Color.BLACK);
				p21.setForeground(Color.BLACK);
				p23.setForeground(Color.BLACK);
				
				p25.setForeground(Color.BLACK);
				
				p28.setForeground(Color.BLACK);
				p31.setForeground(Color.BLACK);
				p34.setForeground(Color.BLACK);
				p37.setForeground(Color.BLACK);
				p40.setForeground(Color.BLACK);
				p43.setForeground(Color.BLACK);
				p46.setForeground(Color.BLACK);
				p49.setForeground(Color.BLACK);
				
				
				
				
				
				value_0 = Double.parseDouble((input_box.getText()));
				try {
					poi();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			System.out.println("control returned");
			}
		});
		btnCalculate.setBounds(105, 488, 132, 25);
		contentPane.add(btnCalculate);
		
		support_2 = new JTextField();
		support_2.setColumns(10);
		support_2.setBounds(154, 565, 55, 20);
		contentPane.add(support_2);
		
		support_3 = new JTextField();
		support_3.setColumns(10);
		support_3.setBounds(210, 585, 55, 20);
		contentPane.add(support_3);
		
		support_4 = new JTextField();
		support_4.setColumns(10);
		support_4.setBounds(263, 606, 55, 20);
		contentPane.add(support_4);
		
		support_5 = new JTextField();
		support_5.setColumns(10);
		support_5.setBounds(319, 624, 55, 20);
		contentPane.add(support_5);
		
		support_0 = new JTextField();
		support_0.setColumns(10);
		support_0.setBounds(39, 524, 55, 20);
		contentPane.add(support_0);
		
		resistance_0 = new JTextField();
		resistance_0.setColumns(10);
		resistance_0.setBounds(39, 462, 55, 20);
		contentPane.add(resistance_0);
		
		resistance_1 = new JTextField();
		resistance_1.setColumns(10);
		resistance_1.setBounds(99, 442, 55, 20);
		contentPane.add(resistance_1);
		
		resistance_2 = new JTextField();
		resistance_2.setColumns(10);
		resistance_2.setBounds(154, 424, 55, 20);
		contentPane.add(resistance_2);
		
		resistance_3 = new JTextField();
		resistance_3.setColumns(10);
		resistance_3.setBounds(210, 404, 55, 20);
		contentPane.add(resistance_3);
		
		resistance_4 = new JTextField();
		resistance_4.setColumns(10);
		resistance_4.setBounds(263, 384, 55, 20);
		contentPane.add(resistance_4);
		
		resistance_5 = new JTextField();
		resistance_5.setColumns(10);
		resistance_5.setBounds(319, 364, 55, 20);
		contentPane.add(resistance_5);
		
		tglbtnHighlight = new JToggleButton("Highlight");
		
		tglbtnHighlight.setBounds(10, 339, 73, 20);
		contentPane.add(tglbtnHighlight);
		
		JLabel label = new JLabel("");
		label.setBounds(259, 445, 115, 99);
		contentPane.add(label);
		
		
		
	}
	
	public void poi() throws IOException{
		
		
		//start of poi 
		
				FileInputStream file_in = new FileInputStream(new File("config.dat"));
				
				XSSFWorkbook workbook = new XSSFWorkbook(file_in);
				
				XSSFSheet sheet = workbook.getSheetAt(0);
				
				try{
					
					Cell cell,sample = null;
					Cell cell_value[] = new Cell[60];
					
					Cell cell_resistance[] = new Cell[6];
					Cell cell_support[]	= new Cell[6];
					
					FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
					
					
					cell = sheet.getRow(3).getCell(3);
					
					 cell.setCellValue(value_0);
					 
					 
				
					
					cell_value[1] = sheet.getRow(8).getCell(4);
					cell_value[2] = sheet.getRow(8).getCell(3);
					cell_value[3] = sheet.getRow(7).getCell(3);
					cell_value[4] = sheet.getRow(7).getCell(4);
					cell_value[5] = sheet.getRow(7).getCell(5);
					cell_value[6] = sheet.getRow(8).getCell(5);
					cell_value[7] = sheet.getRow(9).getCell(5);
					cell_value[8] = sheet.getRow(9).getCell(4);
					cell_value[9] =  sheet.getRow(9).getCell(3);
					
					cell_value[11] = sheet.getRow(8).getCell(2);
					cell_value[13] = sheet.getRow(6).getCell(2);
					cell_value[15] = sheet.getRow(6).getCell(4);
					cell_value[17] = sheet.getRow(6).getCell(6);
					cell_value[19] = sheet.getRow(8).getCell(6);
					cell_value[21] = sheet.getRow(10).getCell(6);
					cell_value[23] = sheet.getRow(10).getCell(4);
					cell_value[25] = sheet.getRow(10).getCell(2);
					
					cell_value[28] = sheet.getRow(8).getCell(1);
					cell_value[31] = sheet.getRow(5).getCell(1);
					cell_value[34] = sheet.getRow(5).getCell(4);
					cell_value[37] = sheet.getRow(5).getCell(7);
					cell_value[40] = sheet.getRow(8).getCell(7);
					cell_value[43] = sheet.getRow(11).getCell(7);
					cell_value[46] = sheet.getRow(11).getCell(4);
					cell_value[49] = sheet.getRow(11).getCell(1);
					
					cell_resistance[0] = sheet.getRow(20).getCell(1);       //acceptable imperfect modification swapped 18,1 to 20,1 in cell_support
					cell_resistance[1] = sheet.getRow(17).getCell(1);
					cell_resistance[2] = sheet.getRow(17).getCell(2);
					cell_resistance[3] = sheet.getRow(17).getCell(3);
					cell_resistance[4] = sheet.getRow(17).getCell(4);
					cell_resistance[5] = sheet.getRow(17).getCell(5);
					
					cell_support[0] = sheet.getRow(18).getCell(1);
					cell_support[1] = sheet.getRow(19).getCell(1);
					cell_support[2]	= sheet.getRow(19).getCell(2);
					cell_support[3] = sheet.getRow(19).getCell(3);
					cell_support[4] = sheet.getRow(19).getCell(4);
					cell_support[5] = sheet.getRow(19).getCell(5);
					
					sample = sheet.getRow(5).getCell(3);
					
					
				    XSSFCell	cell_black[] = new XSSFCell[10];
				    
				    cell_black[1] = sheet.getRow(7).getCell(2);
				    cell_black[2] = sheet.getRow(6).getCell(3);
				    cell_black[3] = sheet.getRow(6).getCell(5);
				    cell_black[4] = sheet.getRow(7).getCell(6);
				    cell_black[5] = sheet.getRow(9).getCell(7);
				    cell_black[6] = sheet.getRow(10).getCell(5);
				    cell_black[7] = sheet.getRow(10).getCell(3);
				    cell_black[8] = sheet.getRow(9).getCell(2);
				    
				    
					
					
					
					
					
					
					
					
					
					
					
					
					 evaluator.evaluateAll();                       //important for formula evaluation 
					 
					{	 
					 p[1]=cell_value[1].getNumericCellValue();
					 p[1] = GUI.round(p[1], 2);
				     p1.setText(Double.toString(p[1]));     
				     
				     p[2]=cell_value[2].getNumericCellValue(); 
				     p[2] = GUI.round(p[2], 2);
				     p2.setText(Double.toString(p[2]));
				     
				     p[3]=cell_value[3].getNumericCellValue(); 
				     p[3] = GUI.round(p[3], 2);
				     p3.setText(Double.toString(p[3]));
				     
				     p[4]=cell_value[4].getNumericCellValue();
				     p[4] = GUI.round(p[4], 2);
				     p4.setText(Double.toString(p[4]));
				     
				     p[5]=cell_value[5].getNumericCellValue(); 
				     p[5] = GUI.round(p[5], 2);
				     p5.setText(Double.toString(p[5]));
				     
				     p[6]=cell_value[6].getNumericCellValue(); 
				     p[6] = GUI.round(p[6], 2);
				     p6.setText(Double.toString(p[6]));
				     
				     p[7]=cell_value[7].getNumericCellValue();
				     p[7] = GUI.round(p[7], 2);
				     p7.setText(Double.toString(p[7]));
				     
				     p[8]=cell_value[8].getNumericCellValue();
				     p[8] = GUI.round(p[8], 2);
				     p8.setText(Double.toString(p[8]));
				     //Okay up
				     p[9]=cell_value[9].getNumericCellValue();
				     p[9] = GUI.round(p[9], 2);
				     p9.setText(Double.toString(p[9]));
				     
				     p[11]=cell_value[11].getNumericCellValue();
				     p[11] = GUI.round(p[11], 2);
				     p11.setText(Double.toString(p[11]));
				     
				     p[13]=cell_value[13].getNumericCellValue();
				     p[13] = GUI.round(p[13], 2);
				     p13.setText(Double.toString(p[13]));
				     
				     p[15]=cell_value[15].getNumericCellValue();
				     p[15] = GUI.round(p[15], 2);
				     p15.setText(Double.toString(p[15]));
				     
				     p[17]=cell_value[17].getNumericCellValue(); 
				     p[17]=GUI.round(p[17], 2);
				     p17.setText(Double.toString(p[17]));
				     
				     p[19]=cell_value[19].getNumericCellValue(); 
				     p[19] = GUI.round(p[19], 2);
				     p19.setText(Double.toString(p[19]));
				     
				     p[21]=cell_value[21].getNumericCellValue();
				     p[21] = GUI.round(p[21], 2);
				     p21.setText(Double.toString(p[21]));
				     
				     p[23]=cell_value[23].getNumericCellValue();
				     p[23] = GUI.round(p[23], 2);
				     p23.setText(Double.toString(p[23]));
				     
				     p[25]=cell_value[25].getNumericCellValue(); 
				     p[25] = GUI.round(p[25], 2);
				     p25.setText(Double.toString(p[25]));
				     
				     p[28]=cell_value[28].getNumericCellValue();
				     p[28] = GUI.round(p[28], 2);
				     p28.setText(Double.toString(p[28]));
				     
				     //start of 3rd block
				     
				     p[28]=GUI.round(cell_value[28].getNumericCellValue(),2); 
				     p28.setText(Double.toString(p[28])); 
					 
					 p[31]=GUI.round(cell_value[31].getNumericCellValue(),2);
					 p31.setText(Double.toString(p[31])); 
					 
					 p[34]=GUI.round(cell_value[34].getNumericCellValue(),2);
					 p34.setText(Double.toString(p[34])); 
					 
					 p[37]=GUI.round(cell_value[37].getNumericCellValue(),2);
					 p37.setText(Double.toString(p[37])); 
					 
					 p[40]=GUI.round(cell_value[40].getNumericCellValue(),2);
					 p40.setText(Double.toString(p[40])); 
					 
					 p[43]=GUI.round(cell_value[43].getNumericCellValue(),2);
					 p43.setText(Double.toString(p[43])); 
					 
					 p[46]=GUI.round(cell_value[46].getNumericCellValue(),2);
					 p46.setText(Double.toString(p[46])); 
					 
					 p[49]=GUI.round(cell_value[49].getNumericCellValue(),2);
					 p49.setText(Double.toString(p[49])); 
					 
					 // black box code
					 
					{
						 
					
					
					
					
					}
					 
					 
					 
					 
					 
					 
					   
					 //Resistance value mapping
					 
					 resistance[0] = GUI.round(cell_resistance[0].getNumericCellValue(), 2);
					 resistance_0.setText(Double.toString(resistance[0]));
					 
					 resistance[1] = GUI.round(cell_resistance[1].getNumericCellValue(), 2);
					 resistance_1.setText(Double.toString(resistance[1]));
					 
					 resistance[2] = GUI.round(cell_resistance[2].getNumericCellValue(), 2);
					 resistance_2.setText(Double.toString(resistance[2]));
					 
					 resistance[3] = GUI.round(cell_resistance[3].getNumericCellValue(), 2);
					 resistance_3.setText(Double.toString(resistance[3]));
					 
					 resistance[4] = GUI.round(cell_resistance[4].getNumericCellValue(), 2);
					 resistance_4.setText(Double.toString(resistance[4]));
					 
					 resistance[5] = GUI.round(cell_resistance[5].getNumericCellValue(), 2);
					 resistance_5.setText(Double.toString(resistance[5]));
					 
					 //Support value mapping
					 
					 support[0] = GUI.round(cell_support[0].getNumericCellValue(), 2);
					 support_0.setText(Double.toString(support[0]));
					 
					 support[1] = GUI.round(cell_support[1].getNumericCellValue(), 2);
					 support_1.setText(Double.toString(support[1]));
					 
					 support[2] = GUI.round(cell_support[2].getNumericCellValue(), 2);
					 support_2.setText(Double.toString(support[2]));
					 
					 support[3] = GUI.round(cell_support[3].getNumericCellValue(), 2);
					 support_3.setText(Double.toString(support[3]));
					 
					 support[4] = GUI.round(cell_support[4].getNumericCellValue(), 2);
					 support_4.setText(Double.toString(support[4]));
					 
					 support[5] = GUI.round(cell_support[5].getNumericCellValue(), 2);
					 support_5.setText(Double.toString(support[5]));
					 
					 
					 
				
				     
					}
					
					
					
					//For black box
					if(value_0>=p[11] && value_0<p[13]){
						black_1.setText(Double.toString(value_0));
					black_1.setBackground(Color.BLACK);
					black_1.setForeground(Color.WHITE);
					}
					else 
					if(value_0>=p[13] && value_0<p[15]){
						black_2.setText(Double.toString(value_0));
						black_2.setBackground(Color.BLACK);
						black_2.setForeground(Color.WHITE);
					}
					else
					if(value_0>=p[15] && value_0<p[17]){
						black_3.setText(Double.toString(value_0));
						black_3.setBackground(Color.BLACK);
						black_3.setForeground(Color.WHITE);
						}
					else
					if(value_0>=p[17] && value_0<p[19]){
						black_4.setText(Double.toString(value_0));
						black_4.setBackground(Color.BLACK);
						black_4.setForeground(Color.WHITE);
						}
					else
					if(value_0>=p[19] && value_0<p[21]){
						black_5.setText(Double.toString(value_0));
						black_5.setBackground(Color.BLACK);
						black_5.setForeground(Color.WHITE);
						}
					else 
						if(value_0>=p[21] && value_0<p[23]){
						black_6.setText(Double.toString(value_0));
						black_6.setBackground(Color.BLACK);
						black_6.setForeground(Color.WHITE);
						}
					else
						if(value_0>=p[23] && value_0<p[25]){
							black_7.setText(Double.toString(value_0));
							black_7.setBackground(Color.BLACK);
							black_7.setForeground(Color.WHITE);
							}
					else 
						if(value_0>=p[9] && value_0<p[11]){
							black_8.setText(Double.toString(value_0));
							black_8.setBackground(Color.BLACK);
							black_8.setForeground(Color.WHITE);
						}
					
				
					
					
			    
				    
				    
				    	
				    	
				    
				    
				    
				   
	
					
					
					file_in.close();
				     
				    FileOutputStream outFile =new FileOutputStream(new File("config.dat"));
				    workbook.write(outFile);
				    outFile.close();
				    
					}
					catch (FileNotFoundException e) {
					    e.printStackTrace();
					} catch (IOException e) {
					    e.printStackTrace();
					}
				
				if(tglbtnHighlight.isSelected()){ // highlight set to true
					int i=2;
					if(p[2]==resistance[0] || p[2]==resistance[1] || p[2]==resistance[3] || p[2]==resistance[4] || p[2]==resistance[5] || p[i]==resistance[2])
						p2.setForeground(Color.RED);
					if(p[2]==support[0] || p[2]==support[1] || p[2]==support[2] || p[2]==support[3] || p[2]==support[4]|| p[2]== support[5] )
						p2.setForeground(Color.BLUE);
					
					i++;
					if(p[3]==resistance[0] || p[3]==resistance[1] || p[3]==resistance[3] || p[3]==resistance[4] || p[3]==resistance[5] || p[i]==resistance[2])
						p3.setForeground(Color.RED);
					if(p[3]==support[0] || p[3]==support[1] || p[3]==support[2] || p[3]==support[3] || p[3]==support[4]|| p[3]== support[5] )
						p3.setForeground(Color.BLUE);
					
					i++;
					if(p[4]==resistance[0] || p[4]==resistance[1] || p[4]==resistance[3] || p[4]==resistance[4] || p[4]==resistance[5] || p[i]==resistance[2])
						p4.setForeground(Color.RED);
					if(p[4]==support[0] || p[4]==support[1] || p[4]==support[2] || p[4]==support[3] || p[4]==support[4]|| p[4]== support[5] )
						p4.setForeground(Color.BLUE);
					
					i++;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p5.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p5.setForeground(Color.BLUE);
					
					i++;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p6.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p6.setForeground(Color.BLUE);
					
					i++;
					if(p[7]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p7.setForeground(Color.RED);
					if(p[7]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p7.setForeground(Color.BLUE);
					
					
					i++;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p8.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p8.setForeground(Color.BLUE);
					
					i++;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p9.setForeground(Color.RED);
					
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p9.setForeground(Color.BLUE);
					
					i=i+2;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p11.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p11.setForeground(Color.BLUE);
					
					i=i+2;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p13.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p13.setForeground(Color.BLUE);
					
					i=i+2;
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p15.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p15.setForeground(Color.BLUE);
					
					i=i+2;
					
					if(p[i]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p17.setForeground(Color.RED);
					if(p[i]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p17.setForeground(Color.BLUE);
					i=i+2;
					if(p[19]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p19.setForeground(Color.RED);
					if(p[19]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p19.setForeground(Color.BLUE);
					i=i+2;
					if(p[21]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p21.setForeground(Color.RED);
					if(p[21]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p21.setForeground(Color.BLUE);
					i=i+2;
					if(p[23]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p23.setForeground(Color.RED);
					if(p[23]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p23.setForeground(Color.BLUE);
					i=i+2;
					if(p[25]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p25.setForeground(Color.RED);
					if(p[25]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p25.setForeground(Color.BLUE);
					i=i+3;
					if(p[28]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p28.setForeground(Color.RED);
					if(p[28]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p28.setForeground(Color.BLUE);
					i=i+3;
					if(p[31]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p31.setForeground(Color.RED);
					if(p[31]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p31.setForeground(Color.BLUE);
					i=i+3;
					if(p[34]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p34.setForeground(Color.RED);
					if(p[34]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p34.setForeground(Color.BLUE);
					i=i+3;
					if(p[37]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p37.setForeground(Color.RED);
					if(p[37]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p37.setForeground(Color.BLUE);
					i=i+3;
					if(p[40]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p40.setForeground(Color.RED);
					if(p[40]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p40.setForeground(Color.BLUE);
					i=i+3;
					if(p[43]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5] || p[i]==resistance[2])
						p43.setForeground(Color.RED);
					if(p[43]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p43.setForeground(Color.BLUE);
					i=i+3;
					if(p[46]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5]|| p[i]==resistance[2] )
						p46.setForeground(Color.RED);
					if(p[46]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p46.setForeground(Color.BLUE);
					i=i+3;
					if(p[49]==resistance[0] || p[i]==resistance[1] || p[i]==resistance[3] || p[i]==resistance[4] || p[i]==resistance[5]|| p[i]==resistance[2] )
						p49.setForeground(Color.RED);
					 if(p[49]==support[0] || p[i]==support[1] || p[i]==support[2] || p[i]==support[3] || p[i]==support[4]|| p[i]== support[5] )
						p49.setForeground(Color.BLUE);
					i=2;
					
					
					
					
				}
				else{System.out.println("2");}
				
				
				
				
				
		}
}

