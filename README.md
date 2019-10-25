# Selenium-Java-Project-
Automation of Export/Import activity using Selenium Java
import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;

import java.awt.Color;

import javax.swing.border.LineBorder;
import javax.swing.JLabel;

import java.awt.Font;

import javax.swing.JComboBox;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JFileChooser;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.DropMode;
import javax.swing.JButton;
import javax.swing.JCheckBox;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.ie.InternetExplorerOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.concurrent.TimeUnit;


public class FinalExport extends JFrame {

	private JPanel contentPane;
	private JPasswordField passwordField_Password;
	private JTextField textField_UserName;
	private JTextField textField_ExportExcelSheet;
	private File filename;
	private String filepath;
	private File foldername;
	private FileInputStream file;
	private XSSFWorkbook workbook;
	private int LastRowCount;
	private int TotalNoOfColumns;
	private String ModellingObjectName;
	private XSSFWorkbook logworkbook;
	private XSSFRow row;
	private int Column1;
	private int Column2;
	private FileOutputStream out;
	private String Revision;
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					FinalExport frame = new FinalExport();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public FinalExport() {
		setBackground(Color.WHITE);
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 584, 465);
		contentPane = new JPanel();
		contentPane.setBackground(Color.WHITE);
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);
		
		JPanel panel = new JPanel();
		panel.setBackground(new Color(211, 211, 211));
		panel.setBorder(new LineBorder(new Color(211, 211, 211), 40));
		contentPane.add(panel, BorderLayout.CENTER);
		panel.setLayout(null);
		
		JPanel panel_1 = new JPanel();
		panel_1.setBorder(new LineBorder(new Color(0, 0, 0), 5));
		panel_1.setBackground(new Color(211, 211, 211));
		panel_1.setBounds(135, 0, 285, 41);
		panel.add(panel_1);
		
		JLabel lblExportEngine = new JLabel("Export Engine - (EE)");
		panel_1.add(lblExportEngine);
		lblExportEngine.setFont(new Font("Tahoma", Font.BOLD, 20));
		
		JPanel panel_2 = new JPanel();
		panel_2.setBackground(new Color(211, 211, 211));
		panel_2.setBorder(new LineBorder(new Color(0, 0, 0), 5));
		panel_2.setBounds(39, 39, 481, 337);
		panel.add(panel_2);
		panel_2.setLayout(null);
		
		JLabel lblLoginPage = new JLabel("LogIn Page");
		lblLoginPage.setBounds(182, 11, 92, 20);
		lblLoginPage.setFont(new Font("Tahoma", Font.BOLD, 16));
		lblLoginPage.setForeground(new Color(0, 0, 0));
		panel_2.add(lblLoginPage);
		
		JLabel lblNewLabel_WebBrowser = new JLabel("Web Browser");
		lblNewLabel_WebBrowser.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_WebBrowser.setBounds(79, 37, 86, 20);
		panel_2.add(lblNewLabel_WebBrowser);
		
		final JComboBox comboBox_WebBrowse = new JComboBox();
		comboBox_WebBrowse.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		comboBox_WebBrowse.setBackground(new Color(255, 255, 255));
		comboBox_WebBrowse.setModel(new DefaultComboBoxModel(new String[] {"Chrome", "Internet Explorer"}));
		comboBox_WebBrowse.setBounds(79, 52, 153, 20);
		panel_2.add(comboBox_WebBrowse);
		
		JLabel lblNewLabel_CamstarPortalName = new JLabel("Camstar Portal Name");
		lblNewLabel_CamstarPortalName.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_CamstarPortalName.setBounds(79, 73, 162, 20);
		panel_2.add(lblNewLabel_CamstarPortalName);
		
		JLabel lblNewLabel_UserName = new JLabel("User Name");
		lblNewLabel_UserName.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_UserName.setBounds(79, 117, 86, 20);
		panel_2.add(lblNewLabel_UserName);
		
		passwordField_Password = new JPasswordField();
		passwordField_Password.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		passwordField_Password.setFont(new Font("Tahoma", Font.BOLD, 11));
		passwordField_Password.setBounds(79, 175, 320, 20);
		panel_2.add(passwordField_Password);
		
		textField_UserName = new JTextField();
		textField_UserName.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		textField_UserName.setFont(new Font("Tahoma", Font.BOLD, 11));
		textField_UserName.setBounds(79, 134, 320, 20);
		panel_2.add(textField_UserName);
		textField_UserName.setColumns(10);
		
		JLabel lblNewLabel_Password = new JLabel("Password");
		lblNewLabel_Password.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_Password.setBounds(79, 160, 86, 14);
		panel_2.add(lblNewLabel_Password);
		
		JLabel lblNewLabel_ExportExcelSheet = new JLabel("Export Excel Sheet");
		lblNewLabel_ExportExcelSheet.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel_ExportExcelSheet.setBounds(56, 224, 135, 20);
		panel_2.add(lblNewLabel_ExportExcelSheet);
		
		final JComboBox comboBox_CamstarPortalName = new JComboBox();
		comboBox_CamstarPortalName.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		comboBox_CamstarPortalName.setEditable(true);
		comboBox_CamstarPortalName.setModel(new DefaultComboBoxModel(new String[] {"https://", "http://"}));
		comboBox_CamstarPortalName.setBounds(79, 93, 320, 20);
		panel_2.add(comboBox_CamstarPortalName);
		
		final JComboBox comboBox_ExportType = new JComboBox();
		comboBox_ExportType.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		comboBox_ExportType.setBackground(new Color(255, 255, 255));
		comboBox_ExportType.setModel(new DefaultComboBoxModel(new String[] {"Single(ROR)", "Multiple(ROR)", "Single(Non ROR)", "Multiple(Non ROR)"}));
		comboBox_ExportType.setBounds(242, 52, 157, 20);
		panel_2.add(comboBox_ExportType);
		
		textField_ExportExcelSheet = new JTextField();
		textField_ExportExcelSheet.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		textField_ExportExcelSheet.setBounds(56, 244, 280, 20);
		panel_2.add(textField_ExportExcelSheet);
		textField_ExportExcelSheet.setColumns(10);
		
		JButton btnNewButton_ExportExcelSheet = new JButton("Browse");
		btnNewButton_ExportExcelSheet.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		btnNewButton_ExportExcelSheet.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				JFileChooser ExportExcelSheet = new JFileChooser();
				ExportExcelSheet.showOpenDialog(null);
				
				filename = ExportExcelSheet.getSelectedFile();
				filepath = filename.getAbsolutePath();
				foldername = ExportExcelSheet.getCurrentDirectory();
				textField_ExportExcelSheet.setText(filepath);	
			}
		});
		btnNewButton_ExportExcelSheet.setBackground(new Color(255, 255, 255));
		btnNewButton_ExportExcelSheet.setFont(new Font("Tahoma", Font.BOLD, 11));
		btnNewButton_ExportExcelSheet.setBounds(346, 243, 89, 23);
		panel_2.add(btnNewButton_ExportExcelSheet);
		
		final JCheckBox chckbxNewCheckBox_ShowPassword = new JCheckBox("Show Password");
		chckbxNewCheckBox_ShowPassword.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(chckbxNewCheckBox_ShowPassword.isSelected()){
					passwordField_Password.setEchoChar((char)0);
				}else{
					passwordField_Password.setEchoChar('*');
				}
				
			}
		});
		chckbxNewCheckBox_ShowPassword.setBackground(new Color(192, 192, 192));
		chckbxNewCheckBox_ShowPassword.setFont(new Font("Tahoma", Font.BOLD, 11));
		chckbxNewCheckBox_ShowPassword.setBounds(79, 197, 135, 20);
		panel_2.add(chckbxNewCheckBox_ShowPassword);
		
		JButton btnNewButton_Export = new JButton("Export");
		btnNewButton_Export.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		btnNewButton_Export.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {

				Object WebBrowser = comboBox_WebBrowse.getSelectedItem();
				String CamstarPortalName = comboBox_CamstarPortalName.getSelectedItem().toString()+".jnj.com/CamstarPortal/Main.aspx";
				String UserName = textField_UserName.getText();
				String Password = passwordField_Password.getText();
				String Domain = "NA";
			    Object ExportType = comboBox_ExportType.getSelectedItem();
			    System.out.println(ExportType);
			    
			    if(WebBrowser.equals("Internet Explorer")){
			    	System.setProperty("webdriver.ie.driver", foldername+"\\IEDriverServer.exe");
			    	WebDriver driver=new InternetExplorerDriver();
			    	driver.manage().window().maximize();
					driver.manage().deleteAllCookies();
					driver.manage().timeouts().pageLoadTimeout(60,TimeUnit.SECONDS);
					driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);
					driver.get(CamstarPortalName);
					WebDriverWait wait = new WebDriverWait(driver,30);
					WebElement webelement1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("UsernameTextbox")));
					webelement1.sendKeys(UserName);
					WebElement webelement2 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PasswordTextbox")));
					webelement2.sendKeys(Password);
					WebElement webelement3 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("DomainDropDown")));
					webelement3.sendKeys(Domain);
					WebElement webelement4 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("LoginButton")));
					webelement4.click();
					File src = new File(filepath);
					
					try {
						file = new FileInputStream(src);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					try {
						workbook = new XSSFWorkbook(file);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					XSSFSheet sheet1 = workbook.getSheetAt(0);
					LastRowCount = sheet1.getLastRowNum();
					TotalNoOfColumns = sheet1.getRow(0).getPhysicalNumberOfCells();
					
					if(ExportType.equals("Single(ROR)")){
					WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
					webelement5.click();
					WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
					webelement6.click();
					
						driver.switchTo().frame(0);
						try {
							Thread.sleep(5000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
						webelement7.click();
						WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						webelement8.click();
						
						for(Column1 = 0;Column1 < TotalNoOfColumns;++Column1){
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							 webelement10.click();
							
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 String RORInstanceName = Instances+"  ROR:"; 
								    System.out.println("RORInstance name = " +Instances);
									WebElement element13A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[starts-with(text(),'"+RORInstanceName+"')] | //div[text() = '"+Instances+"']")));
									element13A.click();
									WebElement element14A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14A.click();
									try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
	                                WebElement element15A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15A.clear();
									
									
								 }
							        WebElement element16A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							        element16A.clear();

									try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
							        
							        
							  }
						WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element17.click();
						try {
							Thread.sleep(1500);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element18.click();
						WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
						element19.click();
						try {
							Thread.sleep(20000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						Date d = new Date();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
						
						File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						try {
							FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						WebElement element20 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_DownloadBtn")));
						element20.click();
						
						
						
					}else if(ExportType.equals("Multiple(ROR)")){
						for(int Column1 = 0;Column1 < TotalNoOfColumns ;++Column1){
							WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
							webelement5.click();
							WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
							webelement6.click();
							int FrameNumber = Column1;
							driver.switchTo().frame(FrameNumber);
							try {
								Thread.sleep(5000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
							webelement7.click();
							WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							webelement8.click();
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 String RORInstanceName = Instances+"  ROR:"; 
								    System.out.println("RORInstance name = " +Instances);
									WebElement element13A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[starts-with(text(),'"+RORInstanceName+"')] | //div[text() = '"+Instances+"']")));
									element13A.click();
									WebElement element14A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14A.click();
									try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
							 
									WebElement element15A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15A.clear();
							
							}
							 WebElement element16A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							 element16A.clear();
							 try {
									Thread.sleep(1500);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							 WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
								element17.click();
								try {
									Thread.sleep(1500);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
								WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
								element18.click();
								WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
								element19.click();
								try {
									Thread.sleep(20000);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
								Date d = new Date();
								SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
								
								File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
								try {
									FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
								} catch (IOException e1) {
									// TODO Auto-generated catch block
									e1.printStackTrace();
								}
								
								
								driver.switchTo().parentFrame();
								try {
									Thread.sleep(2000);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
						
						}	
					}else if(ExportType.equals("Single(Non ROR)")){
						WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
						webelement5.click();
						WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
						webelement6.click();
						driver.switchTo().frame(0);
						try {
							Thread.sleep(5000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
						webelement7.click();
						WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						webelement8.click();
						for(Column1 = 0;Column1 < TotalNoOfColumns;++Column1){
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 
								 int RORColumn1 = (Column1+1);
									System.out.println("RoR Column" +RORColumn1);
									CellType celltype1 = sheet1.getRow(Row).getCell(RORColumn1).getCellType();
									System.out.println("CellType = " +celltype1);
									if(celltype1 == CellType.STRING){
									Revision = sheet1.getRow(Row).getCell(RORColumn1).getStringCellValue();
									}else{
									Revision = NumberToTextConverter.toText(sheet1.getRow(Row).getCell(RORColumn1).getNumericCellValue());	
									}
									System.out.println("Revision = " +Revision);
									String NonRORInstanceName = Instances+":"+Revision; 
									System.out.println("NonRoRInstance name = " +NonRORInstanceName);
									System.out.println("ExportType = " +ExportType);
									WebElement element13B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[text() = '"+NonRORInstanceName+"']")));
									element13B.click();
									WebElement element14B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14B.click();
									try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element15B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15B.clear();
								 
								 
							 }
							 WebElement element16B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
								element16B.clear();
								 try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 Column1 = Column1+1;
								 System.out.println("Last Column Count= " +Column1);
							 
							}
						WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element17.click();
						try {
							Thread.sleep(1500);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element18.click();
						WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
						element19.click();
						try {
							Thread.sleep(20000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						Date d = new Date();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
						
						File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						try {
							FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						WebElement element20 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_DownloadBtn")));
						element20.click();
						
						
					}else{
						for(int Column1 = 0;Column1 < TotalNoOfColumns ;++Column1){
							WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
							webelement5.click();
						    WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
							webelement6.click();
							int FrameNumber = Column1/2;
							driver.switchTo().frame(FrameNumber);
							try {
								Thread.sleep(5000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
							webelement7.click();
							WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							webelement8.click();
							
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 int RORColumn1 = (Column1+1);
									System.out.println("RoR Column" +RORColumn1);
									CellType celltype1 = sheet1.getRow(Row).getCell(RORColumn1).getCellType();
									System.out.println("CellType = " +celltype1);
									if(celltype1 == CellType.STRING){
									Revision = sheet1.getRow(Row).getCell(RORColumn1).getStringCellValue();
									}else{
									Revision = NumberToTextConverter.toText(sheet1.getRow(Row).getCell(RORColumn1).getNumericCellValue());	
									}
									System.out.println("Revision = " +Revision);
									String NonRORInstanceName = Instances+":"+Revision; 
									System.out.println("NonRoRInstance name = " +NonRORInstanceName);
									System.out.println("ExportType = " +ExportType);
									WebElement element13B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[text() = '"+NonRORInstanceName+"']")));
									element13B.click();
									WebElement element14B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14B.click();
									try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element15B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15B.clear();
								 
								 }
							 WebElement element16B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
								element16B.clear();
								 try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 Column1 = Column1+1;
								 System.out.println("Last Column Count= " +Column1);
								 WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
								element17.click();
									try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
									element18.click();
									WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
									element19.click();
									try {
										Thread.sleep(20000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									Date d = new Date();
									SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
									
									File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
									try {
										FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
									} catch (IOException e1) {
										// TODO Auto-generated catch block
										e1.printStackTrace();
									}
									
									driver.switchTo().parentFrame();
									try {
										Thread.sleep(2000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
							 
			
						
						}
	
					}
			    	
			    	
			    	
			    }else{
			    	System.setProperty("webdriver.chrome.driver", foldername+"\\chromedriver.exe");
			    	ChromeOptions options = new ChromeOptions();
					Map<String,Object> prefs = new HashMap<String,Object>();
					prefs.put("profile.default_content_settings.popups",0);
					prefs.put("download.default_directory", foldername+"\\ExportFolder");
					options.setExperimentalOption("prefs",prefs);
					WebDriver driver = new ChromeDriver(options);
					driver.manage().window().maximize();
					driver.manage().deleteAllCookies();
					driver.manage().timeouts().pageLoadTimeout(40,TimeUnit.SECONDS);
					driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
					driver.get(CamstarPortalName);
					WebDriverWait wait = new WebDriverWait(driver,20);
					WebElement webelement1 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("UsernameTextbox")));
					webelement1.sendKeys(UserName);
					WebElement webelement2 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("PasswordTextbox")));
					webelement2.sendKeys(Password);
					WebElement webelement3 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("DomainDropDown")));
					webelement3.sendKeys(Domain);
					WebElement webelement4 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("LoginButton")));
					webelement4.click();
                    File src = new File(filepath);
					
					try {
						file = new FileInputStream(src);
					} catch (FileNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					try {
						workbook = new XSSFWorkbook(file);
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					XSSFSheet sheet1 = workbook.getSheetAt(0);
					LastRowCount = sheet1.getLastRowNum();
					TotalNoOfColumns = sheet1.getRow(0).getPhysicalNumberOfCells();
					
					if(ExportType.equals("Single(ROR)")){
						WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
						webelement5.click();
						WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
						webelement6.click();
						driver.switchTo().frame(0);
						try {
							Thread.sleep(5000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
						webelement7.click();
						WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						webelement8.click();
						for(Column1 = 0;Column1 < TotalNoOfColumns;++Column1){
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							try {
								Thread.sleep(5000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
                                 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 String RORInstanceName = Instances+"  ROR:"; 
								 System.out.println("RORInstance name = " +Instances);
							     WebElement element13A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[starts-with(text(),'"+RORInstanceName+"')] | //div[text() = '"+Instances+"']")));
								 element13A.click();
								 try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 WebElement element14A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
								 element14A.click();
								 try {
										Thread.sleep(2000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 WebElement element15A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
								 element15A.clear();
								 
							 }
							    WebElement element16A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
						        element16A.clear();

								try {
									Thread.sleep(2000);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							
							}
						WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element17.click();
						try {
							Thread.sleep(1500);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element18.click();
						WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
						element19.click();
						try {
							Thread.sleep(20000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						Date d = new Date();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
						
						File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						try {
							FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						WebElement element20 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_DownloadBtn")));
						element20.click();
						
						
						
					}else if(ExportType.equals("Multiple(ROR)")){
						for(int Column1 = 0;Column1 < TotalNoOfColumns ;++Column1){
							WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
							webelement5.click();
							WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
							webelement6.click();
							int FrameNumber = Column1;
							driver.switchTo().frame(FrameNumber);
							try {
								Thread.sleep(5000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
							webelement7.click();
							WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							webelement8.click();
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 try {
										Thread.sleep(500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 String RORInstanceName = Instances+"  ROR:"; 
								 System.out.println("RORInstance name = " +Instances);
								 WebElement element13A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[starts-with(text(),'"+RORInstanceName+"')] | //div[text() = '"+Instances+"']")));
								 element13A.click();
								 WebElement element14A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
								 element14A.click();
								 try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 WebElement element15A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
								 element15A.clear();
								 
							 }
							 WebElement element16A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							 element16A.clear();
							 try {
									Thread.sleep(1500);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							 WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							 element17.click();
							 try {
									Thread.sleep(1500);
								} catch (InterruptedException e) {
									// TODO Auto-generated catch block
									e.printStackTrace();
								}
							 WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							 element18.click();
							 WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
							 element19.click();
							try {
								Thread.sleep(20000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							Date d = new Date();
							SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
							
							File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
							try {
								FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
							} catch (IOException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							driver.switchTo().parentFrame();
							try {
								Thread.sleep(2000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							
						}
						
					}else if(ExportType.equals("Single(Non ROR)")){
						WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
						webelement5.click();
						WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
						webelement6.click();
						driver.switchTo().frame(0);
						try {
							Thread.sleep(5000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
						webelement7.click();
						WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						webelement8.click();
						for(Column1 = 0;Column1 < TotalNoOfColumns;++Column1){
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 int RORColumn1 = (Column1+1);
									System.out.println("RoR Column" +RORColumn1);
									CellType celltype1 = sheet1.getRow(Row).getCell(RORColumn1).getCellType();
									System.out.println("CellType = " +celltype1);
									if(celltype1 == CellType.STRING){
									Revision = sheet1.getRow(Row).getCell(RORColumn1).getStringCellValue();
									}else{
									Revision = NumberToTextConverter.toText(sheet1.getRow(Row).getCell(RORColumn1).getNumericCellValue());	
									}
									System.out.println("Revision = " +Revision);
									String NonRORInstanceName = Instances+":"+Revision; 
									System.out.println("NonRoRInstance name = " +NonRORInstanceName);
									System.out.println("ExportType = " +ExportType);
									WebElement element13B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[text() = '"+NonRORInstanceName+"']")));
									element13B.click();
									WebElement element14B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14B.click();
									try {
										Thread.sleep(1000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element15B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15B.clear();
								 }
							 WebElement element16B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							 element16B.clear();
								 try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 Column1 = Column1+1;
								 System.out.println("Last Column Count= " +Column1);
							}
						WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element17.click();
						try {
							Thread.sleep(1500);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
						element18.click();
						WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
						element19.click();
						try {
							Thread.sleep(20000);
						} catch (InterruptedException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
						Date d = new Date();
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
						
						File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
						try {
							FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
						} catch (IOException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
						
						WebElement element20 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_DownloadBtn")));
						element20.click();
						
					}else{
						for(int Column1 = 0;Column1 < TotalNoOfColumns ;++Column1){
							WebElement webelement5 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("(//a[text() = 'Export/Import'])[1]")));
							webelement5.click();
						    WebElement webelement6 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//li[contains(@class,'navigation-submenuitem')]//a[contains(text(),'Export/Import')]")));
							webelement6.click();
							int FrameNumber = Column1/2;
							driver.switchTo().frame(FrameNumber);
							try {
								Thread.sleep(5000);
							} catch (InterruptedException e) {
								// TODO Auto-generated catch block
								e.printStackTrace();
							}
							WebElement webelement7 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ChoiceWP_ManualExportBtr_ctl00")));
							webelement7.click();
							WebElement webelement8 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
							webelement8.click();
							
							ModellingObjectName = sheet1.getRow(0).getCell(Column1).getStringCellValue();
							WebElement webelement9A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9A.click();
							WebElement webelement9B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							webelement9B.sendKeys(ModellingObjectName);
							WebElement webelement10 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name='ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
							webelement10.click();
							try {
								Thread.sleep(1500);
							} catch (InterruptedException e1) {
								// TODO Auto-generated catch block
								e1.printStackTrace();
							}
							int NumberOfInstances = 1;
							 try{
								    for(NumberOfInstances=1;sheet1.getRow(NumberOfInstances).getCell(Column1).getStringCellValue() != null;++NumberOfInstances){
								    	}
								    
								    }catch(NullPointerException e){
								    	e.printStackTrace();
								    	
								    }
							 System.out.println(NumberOfInstances);
							 for(int Row = 1;Row < NumberOfInstances;Row++){
								 String Instances = sheet1.getRow(Row).getCell(Column1).getStringCellValue();
								 WebElement webelement11A = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11A.click();
								 WebElement webelement11B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//input[@name = 'ctl00$WebPartManager$MDL_Filter_WP$InstanceNameTxt$ctl00']")));
								 webelement11B.sendKeys(Instances);
								 WebElement element12 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//span[contains(@class,'filter-refresh-btn')]")));
								 element12.click();
								 int RORColumn1 = (Column1+1);
									System.out.println("RoR Column" +RORColumn1);
									CellType celltype1 = sheet1.getRow(Row).getCell(RORColumn1).getCellType();
									System.out.println("CellType = " +celltype1);
									if(celltype1 == CellType.STRING){
									Revision = sheet1.getRow(Row).getCell(RORColumn1).getStringCellValue();
									}else{
									Revision = NumberToTextConverter.toText(sheet1.getRow(Row).getCell(RORColumn1).getNumericCellValue());	
									}
									System.out.println("Revision = " +Revision);
									String NonRORInstanceName = Instances+":"+Revision; 
									System.out.println("NonRoRInstance name = " +NonRORInstanceName);
									System.out.println("ExportType = " +ExportType);
									WebElement element13B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//div[text() = '"+NonRORInstanceName+"']")));
									element13B.click();
									WebElement element14B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_SelectBtn")));
									element14B.click();
									try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element15B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Filter_WP_InstanceNameTxt_ctl00")));
									element15B.clear();
								 
							 }
							 WebElement element16B = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_MDL_Objects_WP_ObjectList_Edit")));
							 element16B.clear();
								 try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
								 Column1 = Column1+1;
								 System.out.println("Last Column Count= " +Column1);
								 WebElement element17 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
								element17.click();
									try {
										Thread.sleep(1500);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									WebElement element18 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_NavigationButtonsBar_nextButton")));
									element18.click();
									WebElement element19 = wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("ctl00_WebPartManager_ProgressingWP_StartExportBtn")));
									element19.click();
									try {
										Thread.sleep(20000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
									Date d = new Date();
									SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");
									
									File screenshotFile = ((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
									try {
										FileUtils.copyFile(screenshotFile, new File(foldername+"\\ExportFolder"+"\\"+"Export_"+sdf.format(d)+".png"));
									} catch (IOException e1) {
										// TODO Auto-generated catch block
										e1.printStackTrace();
									}
									
									driver.switchTo().parentFrame();
									try {
										Thread.sleep(2000);
									} catch (InterruptedException e) {
										// TODO Auto-generated catch block
										e.printStackTrace();
									}
							 
							 
							
						}}
			    	
			    	
			    	}	
			
			}});
		btnNewButton_Export.setBackground(new Color(255, 255, 255));
		btnNewButton_Export.setFont(new Font("Tahoma", Font.BOLD, 11));
		btnNewButton_Export.setBounds(105, 285, 86, 23);
		panel_2.add(btnNewButton_Export);
		
		JButton btnNewButton_Reset = new JButton("Reset");
		btnNewButton_Reset.setBorder(new LineBorder(new Color(0, 0, 0), 2));
		btnNewButton_Reset.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				textField_ExportExcelSheet.setText(null);
				textField_UserName.setText(null);
				passwordField_Password.setText(null);
				chckbxNewCheckBox_ShowPassword.setEnabled(false);			
				
			}
		});
		btnNewButton_Reset.setBackground(new Color(255, 255, 255));
		btnNewButton_Reset.setFont(new Font("Tahoma", Font.BOLD, 11));
		btnNewButton_Reset.setBounds(282, 285, 86, 23);
		panel_2.add(btnNewButton_Reset);
		
		JLabel lblNewLabel = new JLabel("Export Type");
		lblNewLabel.setFont(new Font("Tahoma", Font.BOLD, 11));
		lblNewLabel.setBounds(242, 40, 91, 14);
		panel_2.add(lblNewLabel);
		
		
		
		
		
	}
}
