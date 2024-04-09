package test_execution;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.time.Duration;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.Random;

import org.apache.commons.lang3.RandomStringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.StaleElementReferenceException;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.io.FileHandler;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.github.javafaker.Faker;

public class Followup_excel_final_test {
	
	private WebDriver driver;
	private WebElement notify;
	private WebDriverWait wait;
	private JavascriptExecutor js;
	private String date;
	private String PatID;
	private String splcode;
	private String conscode;
	private String consname;
	private String cuscode;
	private String cusname;		
	private double cus_freecon;
	private long cus_freecons;
	private double cus_freecre;
	private long cus_freecred;
	private double cus_singleeven;
	private long cus_singleevent;
	private double cons_freecon;
	private long cons_freecons;
	private double cons_freecre;
	private long cons_freecred;
	private String PM;
	private String Pricename;
	private double PM_rat;
	private long PM_rate;
	private double PM_mindis;
	private long PM_mindisc;
	private double PM_maxdis;
	private long PM_maxdisc;
	private double PM_revisitda;
	private long PM_revisitday;
	private double PM_revisitam;
	private long PM_revisitamt;
	private double C_rat;
	private long C_rate;
	private double C_mindis;
	private long C_mindisc;
	private double C_maxdis;
	private long C_maxdisc;
	private double C_revisitda;
	private long C_revisitday;
	private double C_revisitam;
	private long C_revisitamt;
	private Faker rndname;
	private Random rnd;
	private String Name;
	private int age;
	private String m1;
	private String m2;
	private String id;
	private int maxIterations=8;
	private ExtentReports extnt;
	private ExtentHtmlReporter htmlrprt;
	private ExtentTest testcase;
	private String extntpath="./New_Error_report/extent-report-follownew";
	private String scrnshtpath="./New_Error_report/screenshots/fnew_";
	private String cuspath="./New_Error_report/screenshots/fnew_cusMas";
	private String conspath="./New_Error_report/screenshots/fnew_conpath";
	private String spec_pricepath="./New_Error_report/screenshots/fnew_specpath";
	private String cons_pricepath="./New_Error_report/screenshots/fnew_conpricepath";
	private TakesScreenshot scrnshot;
	private File scrfile;
	private File dstfile;
	private String timestamp;
	private int reff;
	private String min_len;
	private String nation_name;
	private String nation_id;
	private int len;
	private String nation_code;
	private int nat_cod;
	private int rowindex;
	private int rowsh2index;
	private String[] selectvalue4;
	private String[] selectvalue5;
	boolean headerCreated = false;
	boolean headerCreated2 = false;
	private Workbook workbook;
	private Sheet sheet;
	private Sheet sheet1;
	private String Followup_value;
	private String cons_value;
	private int i=1;
	private String result;
	private String cons_result;
	private CellStyle cellcolr;
	private CellStyle cellstyle;
	private String excelfilepath="C:\\Users\\vpm85\\OneDrive\\Desktop\\Followup_result.xlsx";
	private String connect;
	
	private List<String> manfieldname = new ArrayList<>();
	@BeforeClass
	public void setup() {
		rndname=new Faker();
		rnd=new Random();
		timestamp = Long.toString(System.currentTimeMillis() % 10000);
		if(extnt==null) {
			extnt=new ExtentReports();
			htmlrprt=new ExtentHtmlReporter(System.getProperty("user.dir")+extntpath+timestamp+".html");
			extnt.attachReporter(htmlrprt);
			}
		String folderPath = "./New_Error_report/screenshots";
        File folder = new File(folderPath);
        if (!folder.exists()) {
            boolean created = folder.mkdirs();
            System.out.println("Folder is created");
        } else {
            System.out.println("Folder already exists.");
        }
	}
	
	@AfterClass
	public void setupclose() {
		extnt.flush();
	}
	@DataProvider(name ="sqlQueries")
    public Object[][] provideSqlQueries() throws IOException {
        Properties properties = new Properties();
        
        properties.load(Files.newBufferedReader(Paths.get("C:\\Users\\vpm85\\OneDrive\\Desktop\\DB_login.properties")));

        String con=properties.getProperty("connect");
        String usrd=properties.getProperty("userid");
        String pass=properties.getProperty("pasword");
        properties.clear();
        
        properties.load(Files.newBufferedReader(Paths.get("C:\\Users\\vpm85\\OneDrive\\Desktop\\quer_sql.properties")));

        String qy1=properties.getProperty("q1");
        String qy2=properties.getProperty("q2");
        String qy3=properties.getProperty("q3");
        String qy4=properties.getProperty("q4");
        String qy5=properties.getProperty("q5");
        String qy6=properties.getProperty("q6");
        String qy7=properties.getProperty("q7");
        return new Object[][]{
            {con,usrd,pass,qy1,qy2,qy3,qy4,qy5,qy6,qy7}
        };
	}
	
	 @DataProvider(name ="sqlQueries1")
     public Object[][] provideSqlQueries1() throws IOException {
         Properties properties = new Properties();   
     properties.load(Files.newBufferedReader(Paths.get("C:\\Users\\vpm85\\OneDrive\\Desktop\\Webhis_login.properties")));
     
     String chrome=properties.getProperty("chromedriverpath");
     String link=properties.getProperty("webhis_link");
     String usr=properties.getProperty("userid");
     String pas=properties.getProperty("password");

     return new Object[][]{
     				{chrome,link,usr,pas}
     };
	}
	 
	 @Test(dataProvider="sqlQueries")
		public void DB_setup(String DBcon,String DBusr,String DBpas,String qu1,String qu2,String qu3,String qu4,String qu5,String qu6,String qu7 ) throws ClassNotFoundException, SQLException {
			
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			System.out.println("Driver loaded");
			connect=DBcon;
			String user=DBusr;
			String password=DBpas;
			
			Connection con=DriverManager.getConnection(connect, user, password);
			if(con.isClosed()) {
						
			System.out.println("Database is not connected");
					}
			else {

			System.out.println("Database is connected successfully");
					}
			Statement st=con.createStatement();
			
			ResultSet set=st.executeQuery(qu1);
			List<String[]> resultlist=new ArrayList<>();
			while (set.next())
			{
				String[] row=new String[4];
				row[0]=set.getString(1);
				row[1]=set.getString(2);
				row[2]=set.getString(3);
				row[3]=set.getString(4);
				resultlist.add(row);
			}
			String[] selectvalue=resultlist.get(rnd.nextInt(resultlist.size()));
			cuscode=selectvalue[0];
			cusname=selectvalue[1];
			PM=selectvalue[2];
			Pricename=selectvalue[3];
			System.out.println("The Customer code is: "+cuscode);
			System.out.println("The Customer name is: "+cusname);
			System.out.println("The PM code is: "+PM);
			System.out.println("The PM name is: "+Pricename);
			
			ResultSet set1 = st.executeQuery(qu2);
			List<String[]> resultlist1 = new ArrayList<>();

			ResultSetMetaData metaData = set1.getMetaData();
			int numColumns = metaData.getColumnCount();

			while (set1.next()) {
			    String[] row = new String[numColumns];
			    for (int i = 0; i < numColumns; i++) {
			        row[i] = set1.getString(i + 1);
			    }
			    resultlist1.add(row);
			}

			if (!resultlist1.isEmpty()) {
			    String[] selectvalue1 = resultlist1.get(rnd.nextInt(resultlist1.size())); 
			    
			    min_len = selectvalue1[0];
			    len = Integer.parseInt(min_len);
			    nation_id = selectvalue1[1];
			    nation_name = selectvalue1[2];

			    if (selectvalue1.length > 3) {
			        nation_code = selectvalue1[3];
			        nat_cod = Integer.parseInt(nation_code);
			        System.out.println("The nation code is: " + nat_cod);
			    }

			    System.out.println("The min_len is: " + len);
			    System.out.println("The Nation id is: " + nation_id);
			    System.out.println("The Nation name is: " + nation_name);
			}

			
			ResultSet set2 = st.executeQuery(qu3);
			manfieldname = new ArrayList<>(); 
			while (set2.next()) {
			    String ctrlName = set2.getString("Ctrl_Name");
			    manfieldname.add(ctrlName); 
			   }
			System.out.println(manfieldname);
			
			ResultSet set3 = st.executeQuery(qu4);
			List<String[]> resultlist3 = new ArrayList<>();

			ResultSetMetaData metaDataa = set3.getMetaData();
			int numColumn = metaDataa.getColumnCount();

			while (set3.next()) {
			    String[] row = new String[numColumn];
			    for (int i = 0; i < numColumn; i++) {
			        row[i] = set3.getString(i + 1);
			    }
			    resultlist3.add(row);
			}

			if (!resultlist3.isEmpty()) {
			    String[] selectedRow = resultlist3.get(rnd.nextInt(resultlist3.size())); 
			    for (int i = 0; i < numColumn; i++) {
			        //System.out.println(metaDataa.getColumnName(i + 1) + ": " + selectedRow[i]);

			    }  
			    cus_freecon = Double.parseDouble(selectedRow[1]);
				cus_freecons=Math.round(cus_freecon);
				cus_freecre = Double.parseDouble(selectedRow[2]);
				cus_freecred=Math.round(cus_freecre);
				cus_singleeven=Double.parseDouble(selectedRow[3]);
				cus_singleevent=Math.round(cus_singleeven);
				cons_freecon=Double.parseDouble(selectedRow[4]);
				cons_freecons=Math.round(cons_freecon);
				cons_freecre=Double.parseDouble(selectedRow[5]);
				cons_freecred=Math.round(cons_freecre);
				PM_rat=Double.parseDouble(selectedRow[6]);
				PM_rate=Math.round(PM_rat);
				PM_revisitda=Double.parseDouble(selectedRow[7]);
				PM_revisitday=Math.round(PM_revisitda);
				PM_revisitam=Double.parseDouble(selectedRow[8]);
				PM_revisitamt=Math.round(PM_revisitam);
				PM_mindis=Double.parseDouble(selectedRow[9]);
				PM_mindisc=Math.round(PM_mindis);
				PM_maxdis=Double.parseDouble(selectedRow[10]);
				PM_maxdisc=Math.round(PM_maxdis);
				C_rat=Double.parseDouble(selectedRow[11]);
				C_rate=Math.round(C_rat);
				C_mindis=Double.parseDouble(selectedRow[12]);
				C_mindisc=Math.round(C_mindis);
				C_maxdis=Double.parseDouble(selectedRow[13]);
				C_maxdisc=Math.round(C_maxdis);
				C_revisitda=Double.parseDouble(selectedRow[14]);
				C_revisitday=Math.round(C_revisitda);
				C_revisitam=Double.parseDouble(selectedRow[15]);
				C_revisitamt=Math.round(C_revisitam);
				
			}
			
			ResultSet set4=st.executeQuery(qu5);
			List<String[]> resultlist4=new ArrayList<>();
			while (set4.next())
			{
				String[] row=new String[8];
				row[0]=set4.getString(1);
				row[1]=set4.getString(2);
				row[2]=set4.getString(3);
				row[3]=set4.getString(4);
				row[4]=set4.getString(5);
				row[5]=set4.getString(6);
				row[6]=set4.getString(7);
				row[7]=set4.getString(8);
				resultlist4.add(row);
			}
			selectvalue4=resultlist4.get(rnd.nextInt(resultlist4.size()));
			String followupid=selectvalue4[0];
			String day1=selectvalue4[1];
			String day2=selectvalue4[2];
			String day3=selectvalue4[3];
			String day4=selectvalue4[4];
			String day5=selectvalue4[5];
			String day6=selectvalue4[6];
			String day7=selectvalue4[7];
			
			System.out.println("The Followupid is: "+followupid);
			System.out.println("Day 1: "+day1);
			System.out.println("Day 2: "+day2);
			System.out.println("Day 3: "+day3);
			System.out.println("Day 4: "+day4);
			System.out.println("Day 5: "+day5);
			System.out.println("Day 6: "+day6);
			System.out.println("Day 7: "+day7);
			
			ResultSet set5=st.executeQuery(qu6);
			List<String[]> resultlist5=new ArrayList<>();
			while (set5.next())
			{
				String[] row=new String[8];
				row[0]=set5.getString(1);
				row[1]=set5.getString(2);
				row[2]=set5.getString(3);
				row[3]=set5.getString(4);
				row[4]=set5.getString(5);
				row[5]=set5.getString(6);
				row[6]=set5.getString(7);
				row[7]=set5.getString(8);
				resultlist5.add(row);
			}
			selectvalue5=resultlist5.get(rnd.nextInt(resultlist5.size()));
			String Cfollowupid=selectvalue5[0];
			String Cday1=selectvalue5[1];
			String Cday2=selectvalue5[2];
			String Cday3=selectvalue5[3];
			String Cday4=selectvalue5[4];
			String Cday5=selectvalue5[5];
			String Cday6=selectvalue5[6];
			String Cday7=selectvalue5[7];
			
			System.out.println("The Followupid is: "+Cfollowupid);
			System.out.println("Day 1: "+Cday1);
			System.out.println("Day 2: "+Cday2);
			System.out.println("Day 3: "+Cday3);
			System.out.println("Day 4: "+Cday4);
			System.out.println("Day 5: "+Cday5);
			System.out.println("Day 6: "+Cday6);
			System.out.println("Day 7: "+Cday7);
			
			ResultSet set6=st.executeQuery(qu7);
			 List<String[]> resultlist6=new ArrayList<>();
			 while(set6.next()) {
				 String row[]=new String[4];
				 row[0]=set6.getString(1);
				 row[1]=set6.getString(2);
				 row[2]=set6.getString(3);
				 row[3]=set6.getString(4);
				 resultlist6.add(row);
			 }
			 
			String[] selectvalue6= resultlist6.get(rnd.nextInt(resultlist6.size()));
			conscode=selectvalue6[0];
			consname=selectvalue6[1];
			splcode=selectvalue6[3];
			
			System.out.println("splcode is: "+splcode);
			System.out.println("consname is: "+consname);
			System.out.println("conscode is: "+conscode);
			
		}
	 
	 @Test(dataProvider="sqlQueries1")
		public void followup_display(String driverpath,String Webhislink, String usr, String pas) throws InterruptedException, ClassNotFoundException, SQLException, IOException
		{

			Name=rndname.name().firstName();
			age=rnd.nextInt(15, 40);
			reff=rnd.nextInt(001, 026);
			m1=RandomStringUtils.randomNumeric(len);
			m2=RandomStringUtils.randomNumeric(len);
			id=RandomStringUtils.randomNumeric(10);
		
		System.setProperty("webdriver.chrome.driver", driverpath);
		driver=new ChromeDriver();
		driver.get(Webhislink);
		driver.manage().window().maximize();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		js= (JavascriptExecutor)driver;
		driver.findElement(By.id("txtUsrId")).sendKeys(usr);
		driver.findElement(By.id("txtUsrpwd")).sendKeys(pas);
		Thread.sleep(1000);
		driver.findElement(By.id("txtLogin")).click();

		wait=new WebDriverWait(driver,Duration.ofSeconds(20));
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.id("txtLogin")));
		Thread.sleep(3000);
		List<WebElement> menu=driver.findElements(By.xpath("//div[@class='input-group']//ul//li"));
		 if (!menu.isEmpty() && menu.get(0).isDisplayed()) {
	            menu.get(0).click();
	        }
		//wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("(//div[@class='input-group']//ul//li)[1]")));		
		Thread.sleep(5000);
		
		
		WebElement app=driver.findElement(By.id("btnappicon"));
		js.executeScript("arguments[0].click()", app);
		Thread.sleep(2000);
		testcase=extnt.createTest("Navigating to Customer master");
		WebElement mas=driver.findElement(By.xpath("(//span[@id='Masters'])[2]"));
		js.executeScript("arguments[0].click()", mas);
		Thread.sleep(2000);
		WebElement cusmas=driver.findElement(By.xpath("//span[@id='Customers']"));
		js.executeScript("arguments[0].click()", cusmas);
		Thread.sleep(2000);
		WebElement cusr=driver.findElement(By.xpath("//span[text()='Customer']"));
		js.executeScript("arguments[0].click()", cusr);
		driver.findElement(By.xpath("//input[@id='txtCodes']")).click();
		driver.findElement(By.xpath("//input[@id='txtCodes']")).sendKeys(cuscode);
		testcase.log(Status.INFO, "Customer code: "+cuscode);
		testcase.log(Status.INFO, "Customer name: "+cusname);
		Thread.sleep(2000);
		WebElement edit=driver.findElement(By.xpath("//span[@class='glyphicon glyphicon-pencil edit-ico']"));
		List<WebElement> custmr=driver.findElements(By.xpath("//table[@id='tablVirtual']//tbody//tr//td"));
		for(WebElement Customer:custmr) {
			
			 if (Customer.getText().equals(cuscode)) {
			        Thread.sleep(1000);
			        js.executeScript("arguments[0].click()", edit);
			        break;
			    }
			}
		wait.until(ExpectedConditions.visibilityOfAllElements(edit));
		Thread.sleep(1000);
		WebElement cls=driver.findElement(By.xpath("//label[@id='lblClassDetails']"));
		js.executeScript("arguments[0].click()", cls);
		WebElement clsedit=driver.findElement(By.xpath("//table[@class='table table-bordered table-striped table-white']//tbody//tr//td[3]//button"));
		js.executeScript("arguments[0].click()", clsedit);
		Thread.sleep(1000);
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriod']")).click();
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriod']")).clear();
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriod']")).sendKeys(Long.toString(cus_freecons));
		testcase.log(Status.INFO, "Free consultation period for cash in customer master: "+cus_freecons);
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriodforCredit']")).click();
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriodforCredit']")).clear();
		driver.findElement(By.xpath("//input[@id='txtFreeConsultationPeriodforCredit']")).sendKeys(Long.toString(cus_freecred));
		testcase.log(Status.INFO, "Free consultation period for credit in customer master: "+cus_freecred);
		driver.findElement(By.xpath("//input[@id='txtMaximumNumberofDays']")).click();
		driver.findElement(By.xpath("//input[@id='txtMaximumNumberofDays']")).clear();
		driver.findElement(By.xpath("//input[@id='txtMaximumNumberofDays']")).sendKeys(Long.toString(cus_singleevent));
		testcase.log(Status.INFO, "Single event: "+cus_singleevent);
		Thread.sleep(2000);
		String scrpath1=System.getProperty("user.dir")+(cuspath+timestamp+".png");
		scrnshot=(TakesScreenshot)driver;
		scrfile=scrnshot.getScreenshotAs(OutputType.FILE);
		dstfile= new File(scrpath1);
		FileHandler.copy(scrfile, dstfile);
		testcase.addScreenCaptureFromPath(scrpath1);
		driver.findElement(By.xpath("//span[@id='btnSave']")).click();
		Thread.sleep(3000);
		
		WebElement cons=driver.findElement(By.xpath("//span[@id='Consultants']"));
		js.executeScript("arguments[0].click()", cons);
		Thread.sleep(2000);
		WebElement consultan_schedule=driver.findElement(By.xpath("//span[@id='Consultant Schedule']"));
		js.executeScript("arguments[0].click()", consultan_schedule);
		Thread.sleep(1000);
		
		WebElement consul_textbox=driver.findElement(By.xpath("(//input[@id='txtConsultant'])[1]"));
		consul_textbox.sendKeys(conscode);
		Thread.sleep(2000);
		List<WebElement> conslit=driver.findElements(By.xpath("//tbody[@id='ScrollableContent']//tr"));
		for(WebElement con: conslit) {
			con.getText().equals(conscode);
			Thread.sleep(1000);
			con.click();
		}
		Thread.sleep(2000);
		WebElement todate=driver.findElement(By.xpath("//input[@id='txtToDate']"));
		todate.sendKeys(Keys.chord(Keys.CONTROL,"a"));
		todate.sendKeys(Keys.DELETE);
		LocalDate consdate=LocalDate.now().plusMonths(2);
		DateTimeFormatter formatter=DateTimeFormatter.ofPattern("dd/MM/yyyy");
		String condate=consdate.format(formatter);
		todate.sendKeys(condate);
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[@id='btnLoad']")).click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("//button[@id='btnTime']")).click();
		wait.until(ExpectedConditions.presenceOfAllElementsLocatedBy(By.xpath("//button[@id='btnTime']")));	
		WebElement time1=driver.findElement(By.xpath("//input[@id='timfirst1']"));
		time1.click();
		time1.sendKeys("8.00");
		WebElement time2=driver.findElement(By.xpath("//input[@id='timfirst2']"));
		time2.click();
		time2.sendKeys("9.00");
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='chkSelectAll']")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[@id='btnok']")).click();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[@id='btnSave']")).click();
		Thread.sleep(2000);
		testcase=extnt.createTest("Navigating to Consultant master");
		WebElement consl=driver.findElement(By.xpath("//span[@id='Consultant']"));
		js.executeScript("arguments[0].click()", consl);
		Thread.sleep(2000);
		driver.findElement(By.xpath("//input[@id='txtSearch']")).click();
		driver.findElement(By.xpath("//input[@id='txtSearch']")).sendKeys(conscode);
		testcase.log(Status.INFO, "Consultant code: "+conscode);
		WebElement consedit=driver.findElement(By.xpath("//span[@class='glyphicon glyphicon-pencil edit-ico']"));
		List<WebElement> conslist=driver.findElements(By.xpath("//table[@id='data']//tbody//tr//td"));
		for(WebElement consultant:conslist) {
			
			 if (consultant.getText().equals(conscode)) {
			        Thread.sleep(2000);
			        js.executeScript("arguments[0].click()", consedit);
			        break;
			    }
			}
		wait.until(ExpectedConditions.visibilityOfAllElements(consedit));
		Thread.sleep(1000);
		WebElement cons_cash=driver.findElement(By.xpath("//input[@id='txtFreeconslt']"));
		js.executeScript("arguments[0].click()", cons_cash);
		cons_cash.clear();
		cons_cash.sendKeys(Long.toString(cons_freecons));
		testcase.log(Status.INFO, "Free consultation period for cash in consultant master: "+cons_freecons);
		WebElement cons_credit=driver.findElement(By.xpath("//input[@id='txtCreditreg']"));
		js.executeScript("arguments[0].click()", cons_credit);
		cons_credit.clear();
		cons_credit.sendKeys(Long.toString(cons_freecred));
		testcase.log(Status.INFO, "Free consultation period for credit in consultant master: "+cons_freecred);
		Thread.sleep(2000);
		String scrpath2=System.getProperty("user.dir")+(conspath+timestamp+".png");
		scrnshot=(TakesScreenshot)driver;
		scrfile=scrnshot.getScreenshotAs(OutputType.FILE);
		dstfile= new File(scrpath2);
		FileHandler.copy(scrfile, dstfile);
		testcase.addScreenCaptureFromPath(scrpath2);
		driver.findElement(By.xpath("//span[@id='btnSave']")).click();
		Thread.sleep(1000);
		WebElement y=driver.findElement(By.xpath("//button[contains(text(),'Yes')]"));
		js.executeScript("arguments[0].click()", y);
		WebElement emrsave=driver.findElement(By.xpath("//button[contains(text(),'Save')]"));
		js.executeScript("arguments[0].click()", emrsave);
		Thread.sleep(1000);
		driver.findElement(By.xpath("//button[contains(text(),'Yes')]")).click();
		Thread.sleep(1000);
		WebElement save=driver.findElement(By.xpath("(//span[@id='btnSave'])[2]"));
		js.executeScript("arguments[0].click()", save);
		Thread.sleep(1000);
		driver.findElement(By.xpath("//button[contains(text(),'Yes')]")).click();
		Thread.sleep(6000);
		WebElement close = driver.findElement(By.xpath("(//button[@aria-label='Close']//span)[2]"));
		wait.until(ExpectedConditions.elementToBeClickable(close));
		js.executeScript("arguments[0].click()", close);
		
		WebElement priceset=driver.findElement(By.id("Price Settings"));
	    js.executeScript("arguments[0].click()", priceset);
	    testcase=extnt.createTest("Navigating to Speciality price setting");
	    Thread.sleep(2000);
	    WebElement specrate=driver.findElement(By.id("Speciality Rate Setting"));
	    js.executeScript("arguments[0].click()", specrate);
	    wait.until(ExpectedConditions.visibilityOfAllElements(specrate));
	    driver.findElement(By.xpath("(//input[@id='txtTestPrice'])[1]")).clear();
	    driver.findElement(By.xpath("(//input[@id='txtTestPrice'])[1]")).sendKeys(PM);
	    testcase.log(Status.INFO, "Price master: "+PM);
	    Thread.sleep(3000);
	    List<WebElement>pricemas=driver.findElements(By.xpath("//ul[@id='ScrollableContent']//li//div"));
	    for(WebElement pm:pricemas) {
	    	if(pm.getText().equals(PM)) {
	    		Thread.sleep(2000);
	    		js.executeScript("arguments[0].click()", pm);
	    		break;
	    	}
	    }
	    Thread.sleep(1000);
	    driver.findElement(By.id("btnLoad")).click();
	    driver.findElement(By.xpath("//input[@name='lblSearch']")).sendKeys(splcode);
	    Thread.sleep(2000);
	    WebElement checkbox= driver.findElement(By.xpath("(//tbody//tr//td//input)[1]"));
	    if(!checkbox.isSelected())
	    {
	    	checkbox.click();
	    }
	    Thread.sleep(2000);
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[2]")).clear();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[2]")).click();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[2]")).sendKeys(Long.toString(PM_rate));
	    testcase.log(Status.INFO, "Speciality price master rate: "+PM_rate);
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[3]")).clear();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[3]")).click();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[3]")).sendKeys(Long.toString(PM_mindisc));
	    testcase.log(Status.INFO, "Speciality price master minimum discount: "+PM_mindisc);
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[5]")).clear();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[5]")).click();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[5]")).sendKeys(Long.toString(PM_maxdisc));
	    testcase.log(Status.INFO, "Speciality price master maximum discount: "+PM_maxdisc);
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[7]")).clear();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[7]")).click();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[7]")).sendKeys(Long.toString(PM_revisitday));
	    testcase.log(Status.INFO, "Speciality price master revisit day: "+PM_revisitday);
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[8]")).clear();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[8]")).click();
	    driver.findElement(By.xpath("(//tbody//tr//td//input)[8]")).sendKeys(Long.toString(PM_revisitamt));
	    testcase.log(Status.INFO, "Speciality price master revisit amount: "+PM_revisitamt);
	    Thread.sleep(1000);
	    String scrpath3=System.getProperty("user.dir")+(spec_pricepath+timestamp+".png");
		scrnshot=(TakesScreenshot)driver;
		scrfile=scrnshot.getScreenshotAs(OutputType.FILE);
		dstfile= new File(scrpath3);
		FileHandler.copy(scrfile, dstfile);
		testcase.addScreenCaptureFromPath(scrpath3);
	    driver.findElement(By.id("btnSave")).click();
	    Thread.sleep(2000);
	    
	    WebElement conspriceset=driver.findElement(By.xpath("//span[text()='Consultant Wise Price Setting']"));
	    js.executeScript("arguments[0].click()", conspriceset);
	    testcase=extnt.createTest("Navigating to Consultant wise price setting");
	    Thread.sleep(2000);
	    driver.findElement(By.id("txtConsultant")).clear();
	    driver.findElement(By.id("txtConsultant")).sendKeys(conscode);
	    testcase.log(Status.INFO, "Consultant code:"+conscode);
	    Thread.sleep(2000);
	    List<WebElement> conslis=driver.findElements(By.xpath("//tbody[@id='ScrollableContent']//tr//td"));
	    for(WebElement conlis:conslis) {
	    	
	    	if(conlis.getText().equals(conscode)) {
	    		Thread.sleep(2000);
	    		conlis.click();
	    		break;
	    	}
	    }
	    WebElement cspl=driver.findElement(By.xpath("//select[@id='txtConsultationService']"));
	    if(!cspl.isSelected()) {
	    	Select sel=new Select(cspl);
	    	sel.selectByVisibleText(splcode);
	    }
	    
	    driver.findElement(By.id("btnLoad")).click();
	    
	    driver.findElement(By.name("lblSearch")).sendKeys(Pricename);
	    driver.findElement(By.xpath("//tbody//tr//td[5]//input")).clear();
	    driver.findElement(By.xpath("//tbody//tr//td[5]//input")).click();
	    driver.findElement(By.xpath("//tbody//tr//td[5]//input")).sendKeys(Long.toString(C_rate));
	    testcase.log(Status.INFO, "Consultant wise price setting rate: "+C_rate);
	    driver.findElement(By.xpath("//tbody//tr//td[6]//input")).clear();
	    driver.findElement(By.xpath("//tbody//tr//td[6]//input")).click();
	    driver.findElement(By.xpath("//tbody//tr//td[6]//input")).sendKeys(Long.toString(C_mindisc));
	    testcase.log(Status.INFO, "Consultant wise price setting mindisc: "+C_mindisc);
	    driver.findElement(By.xpath("//tbody//tr//td[8]//input")).clear();
	    driver.findElement(By.xpath("//tbody//tr//td[8]//input")).click();
	    driver.findElement(By.xpath("//tbody//tr//td[8]//input")).sendKeys(Long.toString(C_maxdisc));
	    testcase.log(Status.INFO, "Consultant wise price setting maxdisc: "+C_maxdisc);
	    driver.findElement(By.xpath("//tbody//tr//td[10]//input")).clear();
	    driver.findElement(By.xpath("//tbody//tr//td[10]//input")).click();
	    driver.findElement(By.xpath("//tbody//tr//td[10]//input")).sendKeys(Long.toString(C_revisitday));
	    testcase.log(Status.INFO, "Consultant wise price setting revisit day: "+C_revisitday);
	    driver.findElement(By.xpath("//tbody//tr//td[11]//input")).clear();
	    driver.findElement(By.xpath("//tbody//tr//td[11]//input")).click();
	    driver.findElement(By.xpath("//tbody//tr//td[11]//input")).sendKeys(Long.toString(C_revisitamt));
	    testcase.log(Status.INFO, "Consultant wise price setting revisit amount: "+C_revisitamt);
	    Thread.sleep(1000);
	    String scrpath4=System.getProperty("user.dir")+(cons_pricepath+timestamp+".png");
		scrnshot=(TakesScreenshot)driver;
		scrfile=scrnshot.getScreenshotAs(OutputType.FILE);
		dstfile= new File(scrpath4);
		FileHandler.copy(scrfile, dstfile);
		testcase.addScreenCaptureFromPath(scrpath4);
	    driver.findElement(By.id("btnSave")).click();
	    Thread.sleep(2000);
	    
	    WebElement fo=driver.findElement(By.id("Front Office"));
		js.executeScript("arguments[0].click()", fo);
		Thread.sleep(1000);
		WebElement patient=driver.findElement(By.id("Patient"));
		js.executeScript("arguments[0].click()", patient);
		Thread.sleep(1000);
		WebElement patreg=driver.findElement(By.id("Patient Registration"));
		js.executeScript("arguments[0].click()", patreg);
		Thread.sleep(2000);
		 List<WebElement> pops = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
	        if (!pops.isEmpty() && pops.get(0).isDisplayed()) {
	            pops.get(0).click();
	        }
		WebElement newpat=driver.findElement(By.id("tbnToolBarNew"));
		wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("tbnToolBarNew")));
		js.executeScript("arguments[0].click()", newpat);
		Thread.sleep(2000);
		 List<WebElement> pops1 = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
	        if (!pops1.isEmpty() && pops1.get(0).isDisplayed()) {
	            pops1.get(0).click();
	        }
		Thread.sleep(2000);
		WebElement Pat = driver.findElement(By.xpath("//div[@id='divPat_ID']//span"));
		String fullid=Pat.getText();
		PatID=fullid.substring(14);
		System.out.println(PatID);
		Thread.sleep(2000);
		driver.findElement(By.id("txtFirst_Name")).sendKeys(Name);
		
		for (String fieldName : manfieldname) {
	        switch (fieldName) {
	        
	        case "First_Name":
	    		driver.findElement(By.id("txtFirst_Name")).sendKeys(Name);
	    		continue;
	        case "DOB":
	        	continue;
	        case "Age":
	    		driver.findElement(By.id("txtAge")).sendKeys(Integer.toString(age));
	    		continue;
	        case "Gender":
	        	WebElement gn=driver.findElement(By.id("ddlGender"));
	    		List<WebElement> gender=gn.findElements(By.tagName("option"));
	    		int genindx=rnd.nextInt(gender.size());
	    		WebElement gnd=gender.get(genindx);
	    		String g=gnd.getText();
	    		System.out.println(g);
	    		gnd.click();
	    		continue;
	        case "Contact_No":
	        	String cou_code="+"+nat_cod;
	        	WebElement code=driver.findElement(By.id("ddlContact_No"));
	        	code.click();
	        	js.executeScript("arguments[0].scrollIntoView(true);", code);
	        	Thread.sleep(1000);
	        	WebElement dropdownOption = driver.findElement(By.xpath("(//option[@value='"+cou_code+"'])[1]"));
	        	dropdownOption.click();
	    		driver.findElement(By.id("txtContact_No")).sendKeys(m1);
	    		Thread.sleep(2000);
	    		WebElement code2=driver.findElement(By.xpath("(//select[@id='ddlContact_No'])[2]"));
	    		code2.click();
	        	js.executeScript("arguments[0].scrollIntoView(true);", code2);
	        	Thread.sleep(1000);
	        	WebElement dropdownOption1 = driver.findElement(By.xpath("(//option[@value='"+cou_code+"'])[2]"));
	        	dropdownOption1.click();
	    		driver.findElement(By.id("txtContact_No_Sec")).sendKeys(m2);
	    		Thread.sleep(2000);
	    		continue;
	        case "Nationality":
	        	WebElement Area=driver.findElement(By.xpath("//input[@id='txtNationality']"));
	    		js.executeScript("arguments[0].scrollIntoView(true);", Area);
	    		Area.click();
	    		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//input[@id='txtNationality']")));
	    		driver.findElement(By.xpath("//input[@id='txtNationality']")).clear();
	    		List<WebElement> nat=driver.findElements(By.tagName("tr"));
	    		int randrowindex=rnd.nextInt(nat.size());
	    		WebElement nation=nat.get(randrowindex);
	    		String n=nation.getText();
	    		System.out.println(n);
	    		nation.click();
	    		Thread.sleep(1000);
	    		continue;
	        case "IDCardNo":
	    		driver.findElement(By.id("txtIDCardNo")).sendKeys(id);
	    		Thread.sleep(1000);
	    		continue;
	        case "Type":
	        	WebElement idtype=driver.findElement(By.id("ddlType"));
	    		List<WebElement> idty=idtype.findElements(By.tagName("option"));
	    		int idrwindx=rnd.nextInt(idty.size()-1)+1;
	    		WebElement idindex=idty.get(idrwindx);
	    		String i=idindex.getText();
	    		System.out.println(i);
	    		idindex.click();
	    		continue;
	        case "CRCustomer":
	        	WebElement CRCus=driver.findElement(By.id("txtCRCustomer"));
	    		js.executeScript("arguments[0].scrollIntoView(true);", CRCus);
	    		CRCus.click();
	    		driver.findElement(By.id("txtCRCustomer")).sendKeys(cuscode);
	    		Thread.sleep(2000);
	    		List<WebElement>cus=driver.findElements(By.xpath("//tbody[@id='ScrollableContent']//tr"));
	    		for(WebElement customer:cus)
	    		{
	    			customer.equals(cuscode);
	    			Thread.sleep(1000);
	    			customer.click();
	    			
	    		}
	    		Thread.sleep(2000);
	    		List<WebElement> pops5 = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
		        if (!pops5.isEmpty() && pops5.get(0).isDisplayed()) {
		            pops5.get(0).click();
		        }
		        continue;
	        case "CRClass":
	        	continue;
	        case "CRWCompany":
	        	continue;
	        case "OPSpeciality":
	    		continue;
	        case "Consultant":
	        	driver.findElement(By.id("txtConsultant")).click();
	    		driver.findElement(By.id("txtConsultant")).sendKeys(conscode);
	    		Thread.sleep(1000);
	    		List<WebElement>consult=driver.findElements(By.xpath("//tbody[@id='ScrollableContent']//tr"));
	    		for(WebElement conss:consult)
	    		{
	    			conss.equals(conscode);
	    			Thread.sleep(1000);
	    			conss.click();
	    			
	    		}
	    		Thread.sleep(2000);
	    		continue;
	        case "OPConsService":
	        	continue;
	        case "MediaRef":
	        	WebElement reference=driver.findElement(By.id("txtMediaRef"));
			    js.executeScript("arguments[0].click()", reference);
			    reference.clear();
			    reference.sendKeys(Integer.toString(reff));
			    Thread.sleep(2000);

			    List<WebElement> ref = driver.findElements(By.xpath("//tbody[@id='ScrollableContent']//tr"));

			    for (WebElement element : ref) {
			        String text = element.getText();
			        System.out.println("Reference ID: "+text);
			        if (text.contains(Integer.toString(reff))) {
			            element.click();
			            break;
			        }
			    }
			    continue;
	        }
	    }
		WebElement cash_radiobutton=driver.findElement(By.xpath("//div[@id='divOPCash']//input"));
		if(!cash_radiobutton.isSelected()) {
			driver.findElement(By.id("lblInsuranceDetails")).click();
			Thread.sleep(2000);
			LocalDate currentdate=LocalDate.now();
			int randyear=rnd.nextInt(2)+1;
			LocalDate exdate=currentdate.plusYears(randyear);
			DateTimeFormatter formatter1=DateTimeFormatter.ofPattern("dd/MM/yyyy");
			String expirydate=exdate.format(formatter1);
			String apprvldate=currentdate.format(formatter1);
			System.out.println("Insurance expiry date: "+expirydate);
			System.out.println("Insurance approval date: "+apprvldate);
			driver.findElement(By.id("txtINSExpdate")).sendKeys(expirydate);
			driver.findElement(By.id("txtApprovalDate")).sendKeys(apprvldate);
			Thread.sleep(1000);
			driver.findElement(By.id("lblConsulatation")).click();
			}
			driver.findElement(By.id("tbnToolBarSave")).click();
			Thread.sleep(3000);
			driver.findElement(By.xpath("(//div[@class='p-t-b']//button)[1]")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//button[@class='btn bt-Grp gr-btn ok-ico']")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//button[@id='btnClose']")).click();
			Thread.sleep(4000);
			driver.findElement(By.xpath("//button[@id='btnClose']")).click();
			Thread.sleep(4000);
			List<WebElement> pops2 = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
		        if (!pops2.isEmpty() && pops2.get(0).isDisplayed()) {
		            pops2.get(0).click();
		        }
			Thread.sleep(2000);
			
			WebElement newpatt=driver.findElement(By.id("tbnToolBarNew"));
			wait.until(ExpectedConditions.visibilityOfAllElementsLocatedBy(By.id("tbnToolBarNew")));
			js.executeScript("arguments[0].click()", newpatt);
			
			 List<WebElement> pops3 = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
		        if (!pops3.isEmpty() && pops3.get(0).isDisplayed()) {
		            pops3.get(0).click();
		        }

		        LocalDate startDate = LocalDate.now();
		        LocalDate endDate = startDate.plusDays(6);

	        List<String> dates = new ArrayList<>();
	        LocalDate currentDate = startDate;

	        while (!currentDate.isAfter(endDate)) {
	            dates.add(currentDate.format(DateTimeFormatter.ofPattern("dd-MM-yyyy"))); // Corrected date format
	            currentDate = currentDate.plusDays(1);
	        }

	        int headerrowindex=0;
	        int sh2headerrowindex=4;
	        File excelFile = new File(excelfilepath);
	        if (excelFile.exists()) {
	            FileInputStream inputStream = new FileInputStream(excelFile);
	            workbook = new XSSFWorkbook(inputStream);
	            sheet = workbook.getSheet("Result");
	            sheet1= workbook.getSheet("Input");
	            inputStream.close();
	        } else {
	            // Create a new Excel file if it doesn't exist
	            workbook = new XSSFWorkbook();
	            sheet = workbook.createSheet("Result");
	            sheet1= workbook.getSheet("Input");
	        }
	        if (sheet1 == null) {
	            sheet1 = workbook.createSheet("Input");
	        }
	      	rowindex=1;
	      	while(sheet.getRow(rowindex)!=null) {
	      		rowindex++;
	      	}
	      	Row row = sheet.createRow(rowindex);
	      	
	      	rowsh2index=5;
	      	while(sheet1.getRow(rowsh2index)!=null) {
	      		rowsh2index++;
	      	}
	      	Row rowsh2=sheet1.createRow(rowsh2index);
	        
	      	for (String date : dates) {
	        	  Thread.sleep(3000);
	        	  driver.findElement(By.xpath("//input[@id='txtDate']")).click();
		          driver.findElement(By.xpath("//input[@id='txtDate']")).clear();
		          driver.findElement(By.xpath("//input[@id='txtDate']")).sendKeys(date);
		          Thread.sleep(4000);

	            WebElement src = driver.findElement(By.xpath("//button[@id='toolTip']"));
	            js.executeScript("arguments[0].click()", src);
	            Thread.sleep(4000);

	            driver.findElement(By.xpath("//button[@id='btnSearch']")).click();
	            Thread.sleep(3000);

	            wait.until(ExpectedConditions.presenceOfElementLocated(By.xpath("//a[@aria-label='Next page']")));

	            outerLoop: for (int iteration = 1; iteration <= maxIterations; iteration++) {
	                List<WebElement> pat = driver.findElements(By.xpath("//div[@class='table-header-fixed-div_4']//tbody//tr"));

	                for (WebElement search : pat) {
	                    try {
	                        String rowValue = search.getText();
	                        if (rowValue.contains(PatID)) {
	                            Thread.sleep(1000);
	                            search.click();
	                            break outerLoop;
	                        }
	                    } catch (StaleElementReferenceException e) {
	                        System.out.println("Element is stale. Re-finding it.");
	                        break;
	                    }
	                }

	                WebElement nextPageLink = driver.findElement(By.xpath("//a[@aria-label='Next page']"));

	                try {
	                    wait.until(ExpectedConditions.elementToBeClickable(nextPageLink));
	                    nextPageLink.click();
	                } catch (StaleElementReferenceException e) {
	                    System.out.println("Next page link is stale. Re-finding it.");
	                    nextPageLink = driver.findElement(By.xpath("//a[@aria-label='Next page']"));
	                    wait.until(ExpectedConditions.elementToBeClickable(nextPageLink));
	                    nextPageLink.click();
	                }
	            }

	            Thread.sleep(2000);
	            WebElement ok = driver.findElement(By.xpath("//button[@id='lblOK']"));
	            js.executeScript("arguments[0].click()", ok);
	            Thread.sleep(1000);
	            List<WebElement> pops4 = driver.findElements(By.xpath("(//div[@class='swal2-actions']//button)[1]"));
		        if (!pops4.isEmpty() && pops4.get(0).isDisplayed()) {
		            pops4.get(0).click();
		        }
	     
			
		        String[] headers = {"Followup_ID", "ExP_Day1","Act_Day1","ExpCons_day1","ActCons_day1","ExP_Day2","Act_Day2","ExpCons_day2","ActCons_day2",
						"Exp_Day3","Act_Day3","ExpCons_day3","ActCons_day3","Exp_Day4","Act_Day4","ExpCons_day4","ActCons_day4","Exp_Day5","Act_Day5",
						"ExpCons_day5","ActCons_day5","Exp_Day6","Act_Day6","ExpCons_day6","ActCons_day6","Exp_Day7","Act_Day7","ExpCons_day7","ActCons_day7","Patient ID","Status"};
						
						WebElement FollowUp_textbox=driver.findElement(By.xpath("//input[@id='txtFollowUpEvent']"));
						Followup_value=FollowUp_textbox.getAttribute("value");
						System.out.println("The followupvalue: "+Followup_value); 
						
						WebElement Conscharge_textbox=driver.findElement(By.xpath("//input[@id='txtOPConsCharge']"));
						cons_value=Conscharge_textbox.getAttribute("value");
						System.out.println("The cons charge: "+cons_value);   
					    
						cellcolr=workbook.createCellStyle();
					    cellcolr.setFillForegroundColor(IndexedColors.RED.getIndex());
					    cellcolr.setFillPattern(FillPatternType.SOLID_FOREGROUND);

					     if(!headerCreated) {
					     Row headerRow = sheet.createRow(headerrowindex);
					      for (int i = 0; i < headers.length; i++) {
					          Cell cell = headerRow.createCell(i);
					          cell.setCellValue(headers[i]);
					      }
					      headerCreated=true;
					     }
					     row.createCell(0).setCellValue(selectvalue4[0]);
					    
					     for(int i=1;i<selectvalue4.length;i++) {
					    	 int cellindex=4*i-3;
					    	 row.createCell(cellindex).setCellValue(selectvalue4[i]);
					    	}
					    
					     if(!Followup_value.isEmpty()) {
								System.out.println("Yes");
								result="Yes";
								 int cellIndex =4*i-2;
								 Cell cell=row.createCell(cellIndex);
							     cell.setCellValue(result);
							     if(!selectvalue4[i].equals(result)) {
								  cell.setCellStyle(cellcolr);  	
							     }    
							  }
							  else {
								System.out.println("No");
								result="No";
								int cellIndex =4*i-2;
								 Cell cell=row.createCell(cellIndex);
							     cell.setCellValue(result);
							     if(!selectvalue4[i].equals(result)) {
									  cell.setCellStyle(cellcolr);  	
								     } 
							  }
					     for(int i=1;i<selectvalue5.length;i++) {
					    	 int cellindex=4*i-1;
					    	 row.createCell(cellindex).setCellValue(selectvalue5[i]);
					    	}
					     
					     if(!cons_value.equals("0.000") && !cons_value.equals("0.00") ) {
					    	 System.out.println("Yes cons");
					    	 cons_result="Yes Cons";
					    	 int cellindex=4*i;
					    	 Cell cell=row.createCell(cellindex);
					    	 cell.setCellValue(cons_result);
					    	 if(!selectvalue5[i].equals(cons_result)) {
					    		 cell.setCellStyle(cellcolr);
					    	 }
					     }
					     else {
					    	 System.out.println("No cons");
					    	 cons_result="No Cons";
					    	 int cellindex=4*i;
					    	 Cell cell=row.createCell(cellindex);
					    	 cell.setCellValue(cons_result);
					    	 if(!selectvalue5[i].equals(cons_result)) {
					    		 cell.setCellStyle(cellcolr);
					    	 }
					     }
					    
					     row.createCell(29).setCellValue(PatID);

					     try {
					          FileOutputStream outputStream = new FileOutputStream(new File(excelfilepath));
					          workbook.write(outputStream);
					          System.out.println("Data has been written successfully to Excel file.");
					      } catch (IOException e) {
					          e.printStackTrace();
					      }
					      i++;
			      
			      Thread.sleep(3000);
			        String screenshotFileName =System.getProperty("user.dir")+(scrnshtpath+timestamp+ date.replace("-", "") + ".png");
			        testcase=extnt.createTest("Patient on Day "+date);
		            scrnshot=(TakesScreenshot)driver;
		            scrfile=scrnshot.getScreenshotAs(OutputType.FILE);
		    		dstfile= new File(screenshotFileName);
		    		FileHandler.copy(scrfile, dstfile);
		    		testcase.addScreenCaptureFromPath(screenshotFileName);
		    		Thread.sleep(2000);
			      }
	        boolean hasredcolor=false;
	        for(int i=1;i<row.getLastCellNum();i++) {
	        	Cell currentcell=row.getCell(i);
	        	if(currentcell!=null) {
	        		cellstyle=currentcell.getCellStyle();
	        	}
	        	if(cellstyle!=null) {
	        		if(cellstyle.getFillForegroundColor()==IndexedColors.RED.getIndex()) {
	        			hasredcolor=true;
	        			break;
	        			}}}
	        Cell status=row.createCell(30);
	        if(hasredcolor) {
	        	status.setCellValue("Fail");
	        	status.setCellStyle(cellcolr);
	        	}
	        else
	        {
	        	status.setCellValue("Pass");
	        }
	        
	        String[] sh2header= {"Followup_ID","Cus_code","Cus_name","CM_Freecash","CM_Freecredit","CM_singleevent",
	        "Cons_Freecash","Cons_Freecredit","PM_Spec_rate","PM_spec_revisitday","PM_Spec_revisitamt","PM_Spec_Mindis",
	        "PM_Spec_Maxdis","PM_cons_rate","PM_cons_Mindis","PM_cons_Maxdis","PM_cons_revisitdat","PM_cons_revisitamt",
	        "PM_code","PM_name","Cons_code","Cons_name","Spl_Name"};
	        
	        int startIndex = connect.indexOf("//");
	        int semicolonind = connect.indexOf(";", startIndex+1);
	        int	endindex=connect.indexOf(";",semicolonind+1);

	        // Extract the desired substring
	        String DBconnect = connect.substring(startIndex + 2, endindex);
	        Row row1=sheet1.createRow(0);
	        row1.createCell(0).setCellValue("Server & DB:");
	        row1.createCell(1).setCellValue(DBconnect);
	        Row row2=sheet1.createRow(1);
	        row2.createCell(0).setCellValue("Link:");
	        row2.createCell(1).setCellValue(Webhislink);
	        Row row3=sheet1.createRow(2);
	        row3.createCell(0).setCellValue("LinkUserId:");
	        row3.createCell(1).setCellValue(usr);
	        Row row4=sheet1.createRow(3);
	        row4.createCell(0).setCellValue("LinkPassword:");
	        row4.createCell(1).setCellValue(pas);
	       
	        if(!headerCreated2) {
			     Row sh2headerRow = sheet1.createRow(sh2headerrowindex);
			      for (int i = 0; i < sh2header.length; i++) {
			          Cell cell = sh2headerRow.createCell(i);
			          cell.setCellValue(sh2header[i]);
			      }
			      headerCreated2=true;
	        }
	        rowsh2.createCell(0).setCellValue(selectvalue4[0]);
	        rowsh2.createCell(1).setCellValue(cuscode);
	        rowsh2.createCell(2).setCellValue(cusname);
	        rowsh2.createCell(3).setCellValue(cus_freecons);
	        rowsh2.createCell(4).setCellValue(cus_freecred);
	        rowsh2.createCell(5).setCellValue(cus_singleevent);
	        rowsh2.createCell(6).setCellValue(cons_freecons);
	        rowsh2.createCell(7).setCellValue(cons_freecred);
	        rowsh2.createCell(8).setCellValue(PM_rate);
	        rowsh2.createCell(9).setCellValue(PM_revisitday);
	        rowsh2.createCell(10).setCellValue(PM_revisitamt);
	        rowsh2.createCell(11).setCellValue(PM_mindisc);
	        rowsh2.createCell(12).setCellValue(PM_maxdisc);
	        rowsh2.createCell(13).setCellValue(C_rate);
	        rowsh2.createCell(14).setCellValue(C_mindisc);
	        rowsh2.createCell(15).setCellValue(C_maxdisc);
	        rowsh2.createCell(16).setCellValue(C_revisitday);
	        rowsh2.createCell(17).setCellValue(C_revisitamt);
	        rowsh2.createCell(18).setCellValue(PM);
	        rowsh2.createCell(19).setCellValue(Pricename);
	        rowsh2.createCell(20).setCellValue(conscode);
	        rowsh2.createCell(21).setCellValue(consname);
	        rowsh2.createCell(22).setCellValue(splcode);
		      
			    FileOutputStream outputStream = new FileOutputStream(new File(excelfilepath));
			    workbook.write(outputStream);
		        workbook.close();
	}
}
