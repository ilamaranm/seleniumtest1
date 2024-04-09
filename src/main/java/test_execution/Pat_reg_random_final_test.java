package test_execution;

import java.io.File;
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
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;
import com.github.javafaker.Faker;

public class Pat_reg_random_final_test {

	private WebDriver driver;
	private WebElement notify;
	private WebDriverWait wait;
	private JavascriptExecutor js;
	private String date;
	private String PatID;
	private String splcode;
	private String conscode;
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
	private String extntpath="./Error_report/extent-report-follownew";
	private String scrnshtpath="./Error_report/screenshots/fnew_";
	private String cuspath="./Error_report/screenshots/fnew_cusMas";
	private String conspath="./Error_report/screenshots/fnew_conpath";
	private String spec_pricepath="./Error_report/screenshots/fnew_specpath";
	private String cons_pricepath="./Error_report/screenshots/fnew_conpricepath";
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
	
	
	
	private List<String> manfieldname = new ArrayList<>();
	
	@BeforeClass
	public void setup() throws ClassNotFoundException {
		rndname=new Faker();
		rnd=new Random();

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
        String qy5=properties.getProperty("q7");
        return new Object[][]{
            {con,usrd,pass,qy1,qy2,qy3,qy4,qy5}
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
	public void DB_setup(String DBcon,String DBusr,String DBpas,String qu1,String qu2,String qu3,String qu4,String qu6 ) throws ClassNotFoundException, SQLException {
		
		Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		System.out.println("Driver loaded");
		String connect=DBcon;
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
			
			
		    
		   /* System.out.println("cus_freecons:"+cus_freecons);
		    System.out.println("cus_freecred:"+cus_freecred);
		    System.out.println("cus_singleevent:"+cus_singleevent);
		    System.out.println("cons_freecons:"+cons_freecons);
		    System.out.println("cons_freecred:"+cons_freecred);
		    System.out.println("PM_rate:"+PM_rate);
		    System.out.println("PM_revisitday:"+PM_revisitday);
		    System.out.println("PM_revisitamt:"+PM_revisitamt);
		    System.out.println("PM_mindisc:"+PM_mindisc);
		    System.out.println("PM_maxdisc:"+PM_maxdisc);
		    System.out.println("C_rate:"+C_rate);
		    System.out.println("C_mindisc:"+C_mindisc);
		    System.out.println("C_maxdisc:"+C_maxdisc);
		    System.out.println("C_revisitday:"+C_revisitday);
		    System.out.println("C_revisitamt:"+C_revisitamt);*/
			
			ResultSet set4=st.executeQuery(qu6);
			 List<String[]> resultlist6=new ArrayList<>();
			 while(set4.next()) {
				 String row[]=new String[4];
				 row[0]=set4.getString(1);
				 row[1]=set4.getString(4);
				 resultlist6.add(row);
			 }
			 
			String[] selectvalue6= resultlist6.get(rnd.nextInt(resultlist6.size()));
			conscode=selectvalue6[0];
			splcode=selectvalue6[1];
			
			System.out.println("splcode is: "+splcode);
			System.out.println("conscode is: "+conscode);
	    
	}

		    
	}
	
	@Test(dataProvider="sqlQueries1")
	public void followup_display(String driverpath,String Webhislink, String usr, String pas) throws InterruptedException, ClassNotFoundException, SQLException
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
	WebElement mas=driver.findElement(By.xpath("(//span[@id='Masters'])[2]"));
	js.executeScript("arguments[0].click()", mas);
	Thread.sleep(2000);
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
	DateTimeFormatter formatte=DateTimeFormatter.ofPattern("dd/MM/yyyy");
	String condate=consdate.format(formatte);
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
		DateTimeFormatter formatter=DateTimeFormatter.ofPattern("dd/MM/yyyy");
		String expirydate=exdate.format(formatter);
		String apprvldate=currentdate.format(formatter);
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

	}

}
