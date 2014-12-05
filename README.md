
Driver file:
##################################################################################################################
package iGenFramework;


/*-----------------------------------------------------------------
Importing Support Classes
-----------------------------------------------------------------*/
import jxl.*;
import java.io.*;
import java.util.*;


@SuppressWarnings({"unused","rawtypes"})
public class Driver {
	
/*-----------------------------------------------------------------
	Public Variable Declaration Section
-------------------------------------------------------------------*/
public static int curModule;
public static int curTestRow;
public static boolean driverFlag = true;
public static boolean teststatusFlag;
public static String driverpath;
public static String resultpath;
public static String testBrowser;
public static String inputData;
public static String[] testModule;
public static String[][] testmethodCollector;
public static Sheet oDriverSheet;
public static Sheet oModuleSheet;
private static HashMap<String, List<String>>[] testDataCollection;
private static Map<String, String> testEnviromentVariable;


/**
 * @param args
 */
public static void invokeDriver(String callingPackage){
	
	HashMap<String, String> args = new HashMap<String, String>();
	invokeDriver(callingPackage, args);
}

public static void invokeDriver(String callingPackage, HashMap args)
{
	int testModuleCnt;
	int curTestCnt;
	int testCnt;
	String tempTestName;
	String testName;
	String[][] testcaseCollection;
	
	Interface.ShowDialog();
	
	File oFile = new File(inputData);
	if(oFile.exists()){
		
		driverpath = new File(oFile.getParent()).getParent();		
		testEnviromentVariable = Framework.xmlEnvironmentReader(driverpath +"/InputData/");
		try{
			
			jxl.Workbook oBook = Workbook.getWorkbook(oFile);
			oDriverSheet = oBook.getSheet("Driver");
			testDataCollection = Framework.testDataCollection();
			testmethodCollector = new String[testModule.length][oDriverSheet.getColumns()];
			testcaseCollection = TestCaseSelector.ShowDialog(oBook);
			Reporter.createHTMLResultFile("Start", "General");
		
			TestModuleLoop:
			for(testModuleCnt=0;testModuleCnt<testModule.length;testModuleCnt++)
			{
				if(!testModule[testModuleCnt].isEmpty()){
					curModule=testModuleCnt;
					testCnt=1;
					oModuleSheet = oBook.getSheet(testModule[testModuleCnt]);
					System.out.println(testModule[testModuleCnt]);
					Reporter.createHTMLResultFile("ModuleStart", "General");
					
					TestSelectionLoop:
					for(curTestCnt=0;curTestCnt<testcaseCollection[testModuleCnt].length;curTestCnt++)
					{
						
						TestRowLoop:
						for(curTestRow=testCnt;curTestRow<=oModuleSheet.getRows();curTestRow++)
						{
							testName = oModuleSheet.getCell(0,0).getContents();
							tempTestName=Framework.fn_retrieveValue(testName,curTestRow);
							if(new String(tempTestName).equals(testcaseCollection[testModuleCnt][curTestCnt])){
								System.out.println(tempTestName);
								teststatusFlag = true;
								Framework.fn_curTestModuleDriver(curTestRow,callingPackage,args);
								testCnt++;
								System.out.println("\n");
								break TestRowLoop;
							}
						}
					}
					Reporter.createHTMLResultFile("ModuleEnd", "General");
				}
			}
			Reporter.createHTMLResultFile("End", "General");
			
		}catch(Exception e){
			System.out.println(e);
		}
		
	}else{		
		System.out.println("InputDatasheet is not found in \n"+ inputData +"\nExecution ended.");
	}
}

/**
 * @return
 */
public static String testDataCollection(String columnName){	
	return testDataCollection[curModule].get(columnName.toString().trim()).get(curTestRow-1);	
}

/**
 * @return
 */
public static String xmlVariables(String key){
	
	String xmlVariableValue;
	
	if(testEnviromentVariable.containsKey(key)){
		xmlVariableValue = testEnviromentVariable.get(key);
	}else{
		xmlVariableValue = "Variable Not Found in input XML";
	}
	
	return xmlVariableValue;	
}

}
##################################################################################################################
Framework File
##################################################################################################################
package iGenFramework;

/*-----------------------------------------------------------------
Importing Support Classes
-----------------------------------------------------------------*/
import jxl.*;
import java.io.*;
import java.lang.reflect.*;
import java.text.*;
import java.util.*;
import javax.xml.parsers.*;
import org.w3c.dom.*;
import org.xml.sax.*;
import alm.otaclient.*;
import alm.otaclient.events.*;
import com4j.*;
import com4j.util.*;
import com4j.stdole.*;


@SuppressWarnings({"unused","unchecked","rawtypes"})
public class Framework
{
/*-----------------------------------------------------------------
Functions List
1. fn_curTestModuleDriver
2. fn_retrieveValue
-------------------------------------------------------------------*/
	
/**
 * @param curTestRow
 * @param callingPackage
 */
public static void fn_curTestModuleDriver( int curTestRow, String callingPackage, HashMap args ) {
	
	int colCnt;
	int rowCnt;
	int intModule;
	int intModuleCnt = 1;
	int intMethodName;
	boolean executionFlag;
	String tempModule;
	String returnVal;
	String tempColHeader;
	String tempMethodName;
	String tMethodName = null;
	String[] tMethod;
	String tempClass;
	String tempMethod;
	Cell oCell;
	Class<?> cClassName;
	Method sMethod = null;
	
	if(Driver.driverFlag){
		
		for(intModule=0;intModule<Driver.testModule.length;intModule++){
			for(rowCnt=intModuleCnt;rowCnt<Driver.oDriverSheet.getRows();rowCnt++){
				oCell = Driver.oDriverSheet.getCell(0,rowCnt);
				tempModule = oCell.getContents().toString().trim();
				if(new String(tempModule).equalsIgnoreCase(Driver.testModule[intModule])){
					for(colCnt=1;colCnt<=Driver.oDriverSheet.getColumns()-1;colCnt++){
						oCell = Driver.oDriverSheet.getCell(colCnt,rowCnt);
						Driver.testmethodCollector[intModule][colCnt-1] = oCell.getContents();
					}
					intModuleCnt++;
				}
			}
		}
		Driver.driverFlag = false;
	}
	
	Reporter.insetResultTableRow(new String[]{"Test Case Name","Test Case Description"},"Header1", Driver.resultpath);
	Reporter.insetResultTableRow(new String[]{"Test Step Description","Expected Result","Actual Result","Status"},"Header2", Driver.resultpath);

    for(colCnt=1;colCnt<=Driver.oModuleSheet.getColumns()-1;colCnt++)
    {
    	tempColHeader = Driver.oModuleSheet.getCell(colCnt,0).getContents();
    	executionFlag = false;
    	tMethodName = null;
    	if((new String(tempColHeader).contains("Keyword"))){
    		tempMethodName = Driver.oModuleSheet.getCell(colCnt,curTestRow).getContents();
    		
    		switch(tempMethodName){
    		
    		case "Yes":
    			executionFlag = true;
    			if(tempColHeader.length() > 7){
    				tMethodName = Driver.testmethodCollector[Driver.curModule][Integer.parseInt(tempColHeader.substring(7))-1];
    			}
    			break;
    			
    		case "No":
    			executionFlag = false;
    			tMethodName = tempMethodName;
    			break;
    			
    		default:
    			executionFlag = true;
    			if(tempColHeader.toLowerCase() == "keyword" | (!tempMethodName.startsWith("Keyword"))){
    				tMethodName = tempMethodName;
    				if(tempMethodName.startsWith("Keyword")){
    					tMethodName = Driver.testmethodCollector[Driver.curModule][Integer.parseInt(tempMethodName.substring(7))-1];
    				}
    			}else if(tempMethodName.startsWith("Keyword") && isNumeric(tempMethodName.substring(7))){
    				tMethodName = Driver.testmethodCollector[Driver.curModule][Integer.parseInt(tempMethodName.substring(7))-1];
    			}
    			break;
    		}
    	}
    	
    	if(executionFlag){
    		
    		if(!tMethodName.contains(".")){
        		tMethodName = "ApplicationTester." + tMethodName;
        	}
    		
    		tMethod = tMethodName.split("\\.");
			tempClass = callingPackage +"."+tMethod[0];
			tempMethod = tMethod[1];
    		
			try {

				cClassName = Class.forName(tempClass);
				Object obj = cClassName.newInstance();				 
				
				if(!args.isEmpty()){
					try {
						Class<HashMap> argTypes = HashMap.class;
						sMethod = cClassName.getDeclaredMethod(tempMethod,argTypes);
						sMethod.invoke(obj,args);
					} catch (NoSuchMethodException mException) {
						sMethod = cClassName.getDeclaredMethod(tempMethod);
						sMethod.invoke(obj);
					}
				}else if(args.isEmpty()){
					sMethod = cClassName.getDeclaredMethod(tempMethod);
					sMethod.invoke(obj);
				}
				
			} catch (ClassNotFoundException | IllegalAccessException 
					| IllegalArgumentException | InvocationTargetException
					| SecurityException | NoSuchMethodException
					| InstantiationException e) {
				System.out.println(e);
			}

    	}
    }
    
	if(!Driver.teststatusFlag){
		Reporter.teststatus[Driver.curModule][1] = Reporter.teststatus[Driver.curModule][1]+1;
		//connectALM("","Passed");
	}else{
		Reporter.teststatus[Driver.curModule][0] = Reporter.teststatus[Driver.curModule][0]+1;
		//connectALM("","Failed");
	}
	
}
	
	
/**
 * @param colHeader
 * @param curRow
 * @return
 */
public static String fn_retrieveValue( String colHeader, int curRow )
{
	int colCnt;
	String tempHeader;
	String sheetValue = null;
	Cell oCell;
	
	for(colCnt=0;colCnt<=Driver.oModuleSheet.getColumns()-1;colCnt++)
		{
		oCell = Driver.oModuleSheet.getCell(colCnt,0);
		tempHeader = oCell.getContents();
		if(new String(tempHeader).equals(colHeader)){
			oCell = Driver.oModuleSheet.getCell(colCnt,curRow);
			sheetValue = oCell.getContents();
			break;
		}
	}
	return sheetValue;
}

/**
 * @param sheet
 * @param colHeader
 * @param curRow
 * @return
 */
public static String fn_retrieveValue( Sheet sheet, String colHeader, int curRow )
{
	int colCnt;
	String tempHeader;
	String sheetValue = null;
	Cell oCell;
	
	for(colCnt=0;colCnt<=sheet.getColumns()-1;colCnt++)
		{
		oCell = sheet.getCell(colCnt,0);
		tempHeader = oCell.getContents();
		if(new String(tempHeader).equals(colHeader)){
			oCell = sheet.getCell(colCnt,curRow);
			sheetValue = oCell.getContents();
			break;
		}
	}
	return sheetValue;
}

/**
 * @param sheet
 * @param curRow
 * @param curcolm
 * @return
 */
public static String fn_retrieveValue( Sheet sheet, int curRow, int curcolm )
{

	String sheetValue = null;
	Cell oCell;
	
	oCell = sheet.getCell(curcolm,curRow);
	sheetValue = oCell.getContents();

	return sheetValue;
}

/**
 * @param str
 * @return
 */
public static boolean isNumeric(String str)
{
    for (char c : str.toCharArray())
    {
        if (!Character.isDigit(c)) return false;
    }
    return true;
}

public static HashMap<String, List<String>>[] testDataCollection(){
	
	int i;
	String tModule;
	String oModule[] = Driver.testModule;
	HashMap<String, List<String>>[] dataBuilder = new HashMap[4];
	
	File oFile = new File(Driver.inputData);
	jxl.Workbook oBook;
	try {
		oBook = Workbook.getWorkbook(oFile);
		for(i=0;i<oModule.length+1;i++){
			if(i==oModule.length){tModule="Driver";}else{tModule = oModule[i];}				
			dataBuilder[i] = (HashMap<String, List<String>>) DataBuilder(oBook,tModule);
		}
		
		
	} catch (Exception e) {
		System.out.println(e.toString());
	}
	
	return dataBuilder;
}

private static Map<String, List<String>> DataBuilder(Workbook oBook, String oMuldule){
	
	int curTestcols;
	Sheet oSheet;		
	Map<String, List<String>> map;
	
	map = new HashMap<String, List<String>>();
	try {
		oSheet = oBook.getSheet(oMuldule);
		for(curTestcols=0;curTestcols<=oSheet.getColumns()-1;curTestcols++){
			map.put(oSheet.getCell(curTestcols,0).getContents().toString().trim(), DataBuilder(curTestcols, oSheet));
		}
	} catch (Exception e) {
		System.out.println(e.toString());
	}
	
	return map;
}

private static List<String> DataBuilder(final int curTestcols, final Sheet oSheet){
	
	List<String> list;
	int curTestRow;
	
	list = new ArrayList<String>();
	for(curTestRow=1;curTestRow<=oSheet.getRows()-1;curTestRow++){
		list.add(oSheet.getCell(curTestcols,curTestRow).getContents().toString().trim());
	}
	return list;		
}

public static Map<String, String> xmlEnvironmentReader(String folderPath){
	
	Map<String, String> map = new HashMap<String, String>();
	String tempFilePath;
	String varName;
	String varValue;
	File folders;
	File[] fileCollection;
	int intNodes;
	int filesCnt;
	
	folders = new File(folderPath);
	if(folders.exists()){		
		fileCollection = folders.listFiles();
		
		for(filesCnt=0;filesCnt<fileCollection.length;filesCnt++){
			if(fileCollection[filesCnt].isFile()){
				if(fileCollection[filesCnt].getName().toLowerCase().endsWith(".xml")){
					tempFilePath = fileCollection[filesCnt].getAbsolutePath();
					try {
						File xmlVariable = new File(tempFilePath);
						DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();			
						DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
						
						Document doc = dBuilder.parse(xmlVariable);
						doc.getDocumentElement().normalize();
						NodeList nodes = doc.getElementsByTagName("Variable");
			
						for(intNodes=0;intNodes<nodes.getLength();intNodes++) {
							Node node = nodes.item(intNodes);				
							if(node.getNodeType() == Node.ELEMENT_NODE) {
								Element element = (Element) node;
								varName = getValue("Name", element);
								varValue = getValue("Value", element);
								map.put(varName,varValue);
							}
						}
			
					} catch (SAXException | IOException | ParserConfigurationException e) {
						e.printStackTrace();
					}
				}					
			}
		}
	}
	return map;
}

private static String getValue(String tag, Element element) {
	 NodeList nodes = element.getElementsByTagName(tag).item(0).getChildNodes();
	 Node node = (Node) nodes.item(0);
	 return node.getNodeValue();
}

private static void connectALM(String testName,String testStatus, String... fileName) {
	
	boolean attachmentFalg = false;
	File attachmentFile = null;
	String url;
	String username;
	String password;
	String domain;
	String project;
	String testFolder;
	String testSet;
	
	//url of ALM
	url = Driver.xmlVariables("ALMURL");
	//user name for login
	username = Driver.xmlVariables("UID");
	//password for login
	password = Driver.xmlVariables("Pwd");
	//domain
	domain = Driver.xmlVariables("Domain");
	//project
	project = Driver.xmlVariables("Project");
	//folder path in test lab
	testFolder = Driver.xmlVariables("strQCLabPath");
	//test set name
	testSet = Driver.xmlVariables("strQCLabSet");
	
	//Attachment file if any
	if(fileName.length > 0){
		if(fileName[0] != null){
			attachmentFile = new File(fileName[0]);
			attachmentFalg = true;
		}
	}

	if(Driver.xmlVariables("ALMUpdate").equalsIgnoreCase("Yes")){
		try {
			ITDConnection almConnection = ClassFactory.createTDConnection();
			almConnection.initConnectionEx(url);
			almConnection.connectProjectEx(domain, project, username, password);
					
			if(almConnection.connected()) {
				ITestFactory myTestFactory = almConnection.testFactory().queryInterface(ITestFactory.class);
				ITestSetTreeManager tsTreeMgr = almConnection.testSetTreeManager().queryInterface(ITestSetTreeManager.class);
				ITestSetFolder tsFolder = tsTreeMgr.nodeByPath(testFolder).queryInterface(ITestSetFolder.class);
				ITestSetFactory tsFactory = tsFolder.testSetFactory().queryInterface(ITestSetFactory.class);
				
				ITDFilter testSetFilter = tsFactory.filter().queryInterface(ITDFilter.class);
				testSetFilter.filter("CY_CYCLE", "'"+ testSet +"'");		
				IList tsList = tsFactory.newList(testSetFilter.text());
		
				if(tsList.count() > 0) {
					for(Com4jObject  testSetName:tsList) {
				
						ITestSet tsTest = testSetName.queryInterface(ITestSet.class);
						IBaseFactory tsSetFactory = tsTest.tsTestFactory().queryInterface(IBaseFactory.class);
						
						ITDFilter testInstaceCollection = tsSetFactory.filter().queryInterface(ITDFilter.class);
						testInstaceCollection.filter("TS_NAME", "'"+ testName +"'");	
						IList tsTestList = tsSetFactory.newList(testInstaceCollection.text());
						
						if(attachmentFalg){
							IAttachmentFactory attachFactory = tsTest.attachments().queryInterface(IAttachmentFactory.class);
					        IAttachment attach = attachFactory.addItem(attachmentFile.getName()).queryInterface(IAttachment.class);
					        IExtendedStorage extAttach = attach.attachmentStorage().queryInterface(IExtendedStorage.class);
					        extAttach.clientPath(attachmentFile.getParent());  
					        extAttach.save(attachmentFile.getName(), true);					        
					        attach.post();
					        attach.refresh(); 
						} else {
							if(tsTestList.count() > 0) {
								for(Com4jObject tstName:tsTestList) {
	
									ITSTest tTest = tstName.queryInterface(ITSTest.class);
							        IRunFactory runFactory = tTest.runFactory().queryInterface(IRunFactory.class);
							        IRun tRun = runFactory.queryInterface(IRun.class);
	
							        DateFormat dtFmt;        	
							    	dtFmt = new SimpleDateFormat("yyyyMMdd_hhmmss");
							    	Date dt = new Date();
							    	tRun = runFactory.addItem(testName +"_"+ dtFmt.format(dt)).queryInterface(IRun.class);
							    	tRun.copyDesignSteps();
							    	tRun.post();
	
							    	IBaseFactory tStepFactory = tRun.stepFactory().queryInterface(IBaseFactory.class);
							    	IList tStepList = tStepFactory.newList("");
	
							    	if(tStepList.count()>0){
							    		for(Com4jObject tSteps:tStepList){
							    			IStep tStep = tSteps.queryInterface(IStep.class);
							    			tStep.status(testStatus);
							    			tStep.field("ST_ACTUAL", "Step Validation "+ testStatus);
							    			tStep.post();
							    			tStep.refresh();
							    		}
							    	}
							    	tRun.status(testStatus);
							    	tRun.post();
							    	tRun.refresh();
	
							        tTest.post();
							        tTest.refresh();									
								}
							} else {
								System.out.println("Test "+ testName +"not found in Test Set"+ testSet);
							}
						}
					}
				} else {
					System.out.println("Test Set "+ testSet +"not found in Test Lab");
					}	
			} else {
				System.out.println("Unable to connect ALM project");
			}	
			almConnection.disconnectProject();
		} catch (Exception e) {
			System.out.println(e.toString());
		}
	}
}

}
##################################################################################################################
Interface file
##################################################################################################################
package iGenFramework;

/*-----------------------------------------------------------------
Importing Support Classes
-----------------------------------------------------------------*/
import java.awt.*;
import java.awt.event.*;
import java.io.*;

import javax.swing.*;
import javax.swing.GroupLayout.*;
import javax.swing.LayoutStyle.*;

import jxl.*;
import jxl.read.biff.BiffException;


@SuppressWarnings({ "unchecked", "rawtypes" })
public class Interface {

	private JFrame frmInterface;
	private GroupLayout groupLayout;
	private JTextField txtDatatable;
	private JComboBox cmbBrowser;
	private JLabel lblModules;
	private JLabel lblDatatable;
	private JLabel lblBrowser;
	private JList lstModule;
	private JButton btnSubmit;
	private JButton btnReset;
	private JScrollPane scrollPane;
	
	/**
	 * Launch the application.
	 */
	public static void ShowDialog() {
		try {
			Interface jWindow = new Interface();
			jWindow.frmInterface.setVisible(true);
			while(jWindow.frmInterface.isVisible()){}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	/**
	 * Create the application.
	 */
	public Interface() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */

	private void initialize() {
		
		
		
		frmInterface = new JFrame();
		frmInterface.setAlwaysOnTop(true);
		frmInterface.getContentPane().setComponentOrientation(ComponentOrientation.LEFT_TO_RIGHT);
		frmInterface.getContentPane().setFont(new Font("Arial", Font.PLAIN, 11));
		frmInterface.setFont(new Font("Arial", Font.PLAIN, 11));
		frmInterface.setTitle("iGen Framework Interface");
		frmInterface.setBounds(100, 100, 425, 350);
		frmInterface.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		final DefaultListModel lstTestDefault = new DefaultListModel();
		
		lblDatatable = new JLabel("Datatable Location");		
		txtDatatable = new JTextField();
		//txtDatatable.setText("C:\\QTP\\iGen\\iGen Framework\\iGen Framework V2.2\\InputData\\DataSheet_Selenium.xls");
		txtDatatable.setName("iDatatable");
		txtDatatable.setColumns(10);
		txtDatatable.addFocusListener(new FocusListener() {

			@Override
			public void focusLost(FocusEvent e) {
				if (txtDatatable.getText().toString().length()!=0){
		        	txtDatatable.setEnabled(false);
		        	btnSubmit.setEnabled(false);
		        	btnReset.setEnabled(false);
		        	String tempPath = txtDatatable.getText().trim();
		        	try {
		        		File oFile = new File(tempPath);
		        		lstTestDefault.removeAllElements();
		        		if(oFile.exists()){
							jxl.Workbook oBook = Workbook.getWorkbook(oFile);
							String[] tempSheetNames = oBook.getSheetNames();
							for(int intI=0;intI<tempSheetNames.length;intI++){
								if(!tempSheetNames[intI].equalsIgnoreCase("Driver")){
									lstTestDefault.addElement(tempSheetNames[intI].toString().trim());
								}
							}
							txtDatatable.setBorder(BorderFactory.createLineBorder(Color.decode("#69BE28")));
							lstModule.setBorder(BorderFactory.createLineBorder(Color.decode("#4B4B4B")));
							oBook.close();
		        		}else{
		        			txtDatatable.setBorder(BorderFactory.createLineBorder(Color.decode("#E65032")));
		        			lstModule.setBorder(BorderFactory.createLineBorder(Color.decode("#E65032")));
		        			lstTestDefault.addElement("Please enter a");
		        			lstTestDefault.addElement("valid Datatable in");
		        			lstTestDefault.addElement("Datatable textfield");
		        		}
					} catch (BiffException | IOException e1) {
						// TODO Auto-generated catch block
						e1.printStackTrace();
					}
		        	txtDatatable.setEnabled(true);
		        	btnSubmit.setEnabled(true);
		        	btnReset.setEnabled(true);
		        }else{
		        	lstTestDefault.removeAllElements();
		        }				
			}

			@Override
			public void focusGained(FocusEvent e) {
				btnSubmit.setEnabled(false);
	        	btnReset.setEnabled(false);
				lstTestDefault.removeAllElements();
				lstTestDefault.addElement("...");
				txtDatatable.setBorder(BorderFactory.createLineBorder(Color.decode("#00B2EF")));
				btnSubmit.setEnabled(true);
	        	btnReset.setEnabled(true);
			}
		});
		
		lblBrowser = new JLabel("Select Explorer to perform Test");		
		cmbBrowser = new JComboBox();
		cmbBrowser.setName("iBrowser");
		cmbBrowser.setModel(new DefaultComboBoxModel(new String[] {"Internet Explorer", "Mozilla Firefox", "Google Chrome", "Apple Safari"}));
		
		lblModules = new JLabel("Select Modules to Perform Test");
		lblModules.setAutoscrolls(true);
		lblModules.setToolTipText("Select Modules to Perform Test");
		
		btnSubmit = new JButton("Submit");
		btnSubmit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				Driver.inputData = txtDatatable.getText().trim();
				Driver.testBrowser = cmbBrowser.getSelectedItem().toString().trim();
				System.out.println(Driver.inputData);
				System.out.println(Driver.testBrowser);
				if(lstModule.getSelectedIndex() != -1){
					int index[] = lstModule.getSelectedIndices();
					Driver.testModule = new String[index.length];
					for(int i=0;i<index.length;i++){
						Driver.testModule[i] = lstModule.getModel().getElementAt(index[i]).toString();
						System.out.println(Driver.testModule[i]);
					}
				}else{
					Driver.testModule = new String[1];
					Driver.testModule[0] = "";
				}
				frmInterface.dispose();
			}
		});
		
		btnReset = new JButton("Reset");
		btnReset.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frmInterface.validate();
				frmInterface.repaint();
			}
		});		
		
		scrollPane = new JScrollPane();
		
		groupLayout = new GroupLayout(frmInterface.getContentPane());
		groupLayout.setHorizontalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGap(38)
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addComponent(lblBrowser, GroupLayout.PREFERRED_SIZE, 160, GroupLayout.PREFERRED_SIZE)
						.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING, false)
							.addComponent(btnSubmit, GroupLayout.PREFERRED_SIZE, 81, GroupLayout.PREFERRED_SIZE)
							.addComponent(lblModules, Alignment.LEADING, GroupLayout.PREFERRED_SIZE, 139, GroupLayout.PREFERRED_SIZE))
						.addComponent(lblDatatable))
					.addGap(18)
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addComponent(scrollPane, 0, 0, Short.MAX_VALUE)
						.addGroup(groupLayout.createParallelGroup(Alignment.LEADING, false)
							.addComponent(txtDatatable, GroupLayout.DEFAULT_SIZE, 153, Short.MAX_VALUE)
							.addComponent(cmbBrowser, 0, 153, Short.MAX_VALUE)
							.addComponent(btnReset, GroupLayout.PREFERRED_SIZE, 86, GroupLayout.PREFERRED_SIZE)))
					.addGap(352))
		);
		groupLayout.setVerticalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGap(55)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblBrowser, GroupLayout.PREFERRED_SIZE, 26, GroupLayout.PREFERRED_SIZE)
						.addComponent(cmbBrowser, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(lblDatatable, GroupLayout.PREFERRED_SIZE, 23, GroupLayout.PREFERRED_SIZE)
						.addComponent(txtDatatable, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE))
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(57)
							.addComponent(lblModules, GroupLayout.PREFERRED_SIZE, 31, GroupLayout.PREFERRED_SIZE))
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(30)
							.addComponent(scrollPane, GroupLayout.PREFERRED_SIZE, 84, GroupLayout.PREFERRED_SIZE)))
					.addGap(32)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnSubmit)
						.addComponent(btnReset))
					.addGap(32))
		);
		lstModule = new JList();
		scrollPane.setViewportView(lstModule);
		lstModule.setModel(lstTestDefault);
		lstModule.setName("iModule");
		frmInterface.getContentPane().setLayout(groupLayout);
	}
}

##################################################################################################################
reporter file
##################################################################################################################
package iGenFramework;

import java.io.*;
import java.text.*;
import java.util.*;

public class Reporter {
	
	public static int[][] teststatus;
	
/**
 * @param strFlowStatus
 * @param strResName
 */
public static void createHTMLResultFile( String strFlowStatus,  String strResName ){
		
	String insertLine = "";
	String input="";
	String line="";
	String strTemp="";
	byte[] contentInBytes;
	DateFormat dtFmt;

	dtFmt = new SimpleDateFormat("dd_MM_yyyy_HH_mm_ss");
	Date dt = new Date();		
	if (strFlowStatus=="Start"){	
		teststatus = new int[Driver.testModule.length][4];
		for(int intI=0;intI<Driver.testModule.length;intI++){
			for(int intJ=0;intJ<teststatus[intI].length;intJ++){
				teststatus[intI][intJ]=0;
			}
		}
		Driver.resultpath = Driver.driverpath +"\\OutputFiles\\"+ strResName +"_"+ dtFmt.format(dt) +".html";
		try {
			File file = new File(Driver.resultpath);
			file.createNewFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	try (FileOutputStream oFso = new FileOutputStream(new File(Driver.resultpath),true)) {
		if (strFlowStatus.equalsIgnoreCase("start")){

			insertLine = insertLine +  "\n <!DOCTYPE html PUBLIC -//W3C//DTD XHTML 1.0 Strict//EN";
			insertLine = insertLine +  "\n http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd> ";
			insertLine = insertLine +  "\n <HTML><HEAD>";
			insertLine = insertLine +  "\n <title>iGen Automation Test Report</title>";
			insertLine = htmlResultFileCSS( insertLine );
			insertLine = insertLine +  "\n </HEAD><BODY>";
			insertLine = insertLine +  "\n <P><H4><CENTER><FONT face=Verdana   size=4 color=#ffffff>iGen Automation Test Report</FONT></CENTER></H4></P>";
			insertLine = insertLine +  "\n <div id=\"tabbed_box_1\" class=\"tabbed_box\">";
			insertLine = insertLine +  "\n <div class=\"tabbed_area\">";
			insertLine = insertLine +  "\n <ul class=\"tabs\">";

			insertLine = insertLine +  "\n <script type=\"text/javascript\"></script>";
			insertLine = insertLine +  "\n <li><a href=\"javascript:tabSwitch('tab_1', 'content_1');\" id=\"tab_1\" class=\"active\"> Summary </a></li>";
			for(int intI=0;intI<Driver.testModule.length;intI++){
				if(!Driver.testModule[intI].isEmpty()){
					insertLine = insertLine +  "\n <li><a href=\"javascript:tabSwitch('tab_"+ (intI+2) +"', 'content_"+ (intI+2) +"');\" id=\"tab_"+ (intI+2) +"\">"+ Driver.testModule[intI] +"</a></li>";
				}
			}
			insertLine = insertLine +  "\n </ul>";
			insertLine = insertLine +  "\n <div id= \"content_1\" class=\"content\">";
			insertLine = insertLine +  "\n <ul>";
			insertLine = insertLine +  "\n <li><TABLE  width='100%' >";
			insertLine = insertLine +  "\n <P><B><U><Center><FONT face=Verdana color=#0033CC  size=4>Execution Status</FONT></Center></U></B></P>";
			insertLine = insertLine +  "\n </Table></li>";
			insertLine = insertLine +  "\n <li><TABLE align=center valign=center border=0 cellpadding=1 cellspacing=1 width= \"100%\" height= \"100%\">";
			insertLine = insertLine +  "\n <TR><TD width= \"100%\" height= \"100%\">";
			insertLine = insertLine +  "\n <div id=\"chart_summary\" style=\"width: 1200px; height: 500px;\">";
			/*
			insertLine = insertLine +  "\n <!Template name=\"Summary\">";

			for(int intI=0;intI<Driver.testModule.length;intI++){
				if(!Driver.testModule[intI].isEmpty()){
					insertLine = insertLine +  "\n <input id=\"module"+ (intI+1) +"testpass"+ (intI+1) +"\" value="+ teststatus[intI][0] +" />";
					insertLine = insertLine +  "\n <input id=\"module"+ (intI+1) +"testfail"+ (intI+1) +"\" value="+ teststatus[intI][1] +" />";
					insertLine = insertLine +  "\n <input id=\"module"+ (intI+1) +"validationpass"+ (intI+1) +"\" value="+ teststatus[intI][2] +" />";
					insertLine = insertLine +  "\n <input id=\"module"+ (intI+1) +"validationfail"+ (intI+1) +"\" value="+ teststatus[intI][3] +" />";
				}
			}
			insertLine = insertLine +  "\n <!/Template>";
			*/
			insertLine = insertLine +  "\n </div></TD></TR>";
			insertLine = insertLine +  "\n </Table></li></ul>";
			insertLine = insertLine +  "\n </div>";

			for(int intI=0;intI<Driver.testModule.length;intI++){
				dtFmt = new SimpleDateFormat("yyyy/MM/dd");	
				insertLine = insertLine +  "\n <div id= \"content_"+ (intI+2) +"\" class=\"content\">";
				insertLine = insertLine +  "\n <ul>";
				insertLine = insertLine +  "\n <li><TABLE border=0 cellSpacing=1 cellPadding=3 width=\"100%\">";
				insertLine = insertLine +  "\n <TBody><TR>";
				insertLine = insertLine +  "\n <TD bgColor=white width=\"6%\">";
				insertLine = insertLine +  "\n <IMG src=\"http://mysource.bnymellon.net/nav/img/tbnymlogo.gif\"></TD>";
				insertLine = insertLine +  "\n <TD><TABLE height=10 width=\"100%\" border=0 cellSpacing=1 cellPadding=3>";
				insertLine = insertLine +  "\n <TBody>";
				insertLine = insertLine +  "\n <TR>";
				insertLine = insertLine +  "\n <TD vAlign=center  align=left  width='3%' height=20>";
				insertLine = insertLine +  "\n <FONT face=Verdana color=#0033CC size=2><B>Execution Date :  ";
				insertLine = insertLine +  dtFmt.format(dt) +"</B></FONT>"; 
				insertLine = insertLine +  "\n <BR><FONT face=Verdana color=#0033CC size=2><B>Executed By : ";
				insertLine = insertLine +  ""+ System.getProperty("user.name");
				insertLine = insertLine +  "\n </B></FONT></TD>";
				insertLine = insertLine +  "\n <TD vAlign=center  align=left  width='3%' height=20>";
				insertLine = insertLine +  "\n <FONT face=Verdana color=#0033CC size=2><B>Tool Name: ";
				insertLine = insertLine +  "Java - "+ System.getProperty("java.vendor");
				insertLine = insertLine + "<BR></FONT>";
				insertLine = insertLine +  "\n <FONT face=Verdana color=#0033CC size=2><B>Tool Version: ";
				insertLine = insertLine + System.getProperty("java.version");
				insertLine = insertLine +  "\n </B></FONT></TD></TR>";
				insertLine = insertLine +  "\n </TBody></Table>";
				insertLine = insertLine +  "\n </TD></TR>";
				insertLine = insertLine +  "\n <!PerformLineInsertModule"+ (intI+1) +">";				
				insertLine = insertLine +  "\n </TBody></Table></li>";
				insertLine = insertLine +  "\n <li><TABLE  width='100%' >";
				insertLine = insertLine +  "\n <P><B><U><Center><FONT face=Verdana color=#0033CC  size=4>Result</FONT></Centert></U></B></P>";
				insertLine = insertLine +  "\n </TABLE>";
				insertLine = insertLine +  "\n <TABLE align=center valign=center border=0 cellpadding=1 cellspacing=1>";
				insertLine = insertLine +  "\n <TR><TD align=center valign=center>";
				insertLine = insertLine +  "\n <div id=\"chart_module"+ (intI+1) +"\" style=\"width: 1200px; height: 200px;\">";
				insertLine = insertLine +  "\n <Template name=\"ModuleName"+ (intI+1) +"\">";
				insertLine = insertLine +  "\n <input id=\"testpass"+ (intI+1) +"\" value=0 />";
				insertLine = insertLine +  "\n <input id=\"testfail"+ (intI+1) +"\" value=0 />";
				insertLine = insertLine +  "\n <input id=\"validationpass"+ (intI+1) +"\" value=0 />";
				insertLine = insertLine +  "\n <input id=\"validationfail"+ (intI+1) +"\" value=0 />";
				insertLine = insertLine +  "\n </Template>";
				insertLine = insertLine +  "\n </div> </TD></TR>";
				insertLine = insertLine +  "\n </Table></li></ul>";
				insertLine = insertLine +  "\n </div>";
			}

			insertLine = insertLine +  "\n </div>";
			insertLine = insertLine +  "\n </div>";
			insertLine = insertLine +  "\n </body>";
			insertLine = insertLine +  "\n <!PerformTestCaseCountChanges><script></script>";
			insertLine = insertLine +  "\n </html>";

			contentInBytes = insertLine.getBytes();
			oFso.write(contentInBytes);
			oFso.flush();
			oFso.close();	
		}else if (strFlowStatus.equalsIgnoreCase("moduleend")){
			try (BufferedReader oFsoRead = new BufferedReader(new FileReader(Driver.resultpath));) {
				while ((line = oFsoRead.readLine()) != null){
					if(line.contains("<!PerformTestCaseCountChanges>")){
						line = " <!PerformTestCaseCountChanges><script>";
						for(int intI=0;intI<Driver.testModule.length;intI++){
							strTemp = strTemp + "document.getElementById('testpass"+ (intI+1) +"').value="+ teststatus[intI][0] +";";
							strTemp = strTemp + "document.getElementById('testfail"+ (intI+1) +"').value="+ teststatus[intI][1] +";";
							strTemp = strTemp + "document.getElementById('validationpass"+ (intI+1) +"').value="+ teststatus[intI][2] +";";
							strTemp = strTemp + "document.getElementById('validationfail"+ (intI+1) +"').value="+ teststatus[intI][3] +";";
						}
						System.out.println("Test");
						line = line + strTemp + "</script>";
					}
					input = input + line + '\n';
				}
				oFsoRead.close();

			}catch (IOException e) {
				e.printStackTrace();
			}

			try (FileOutputStream oFsoWrite = new FileOutputStream(new File(Driver.resultpath),false)) {
				contentInBytes = input.getBytes();
				oFsoWrite.write(contentInBytes);
				oFsoWrite.flush();
				oFsoWrite.close();	
			} catch (IOException e) {
				e.printStackTrace();
			}			
		}
	} catch (IOException e) {
		e.printStackTrace();
	}

}

/**
 * @param insertLine
 * @return
 */
private static String htmlResultFileCSS( String insertLine ){

	insertLine = insertLine +  "\n <Style>";
	insertLine = insertLine +  "\n body{";
	insertLine = insertLine +  "\n background-image:url(images/background.jpg);";
	insertLine = insertLine +  "\n background-repeat:no-repeat;";
	insertLine = insertLine +  "\n background-position:top center;";
	insertLine = insertLine +  "\n background-color:#4B4B4B;";
	insertLine = insertLine +  "\n margin:40px;}";
	insertLine = insertLine +  "\n #tabbed_box{";
	insertLine = insertLine +  "\n margin: 0px auto 0px auto;";
	insertLine = insertLine +  "\n width:300px;}";
	insertLine = insertLine +  "\n .tabbed_box h4 {";
	insertLine = insertLine +  "\n font-family:Arial, Helvetica, sans-serif;";
	insertLine = insertLine +  "\n font-size:23px;";
	insertLine = insertLine +  "\n color:#ffffff;";
	insertLine = insertLine +  "\n letter-spacing:-1px;";
	insertLine = insertLine +  "\n margin-bottom:10px;}";
	insertLine = insertLine +  "\n .tabbed_box h4 small{";
	insertLine = insertLine +  "\n color:#e3e9ec;";
	insertLine = insertLine +  "\n font-weight:normal;";
	insertLine = insertLine +  "\n font-size:9px;";
	insertLine = insertLine +  "\n font-family:Verdana, Arial, Helvetica, sans-serif;";
	insertLine = insertLine +  "\n text-transform:uppercase;";
	insertLine = insertLine +  "\n position:relative;";
	insertLine = insertLine +  "\n top:-4px;";
	insertLine = insertLine +  "\n left:6px;";
	insertLine = insertLine +  "\n letter-spacing:0px;}";
	insertLine = insertLine +  "\n .tabbed_area{";
	insertLine = insertLine +  "\n border:1px solid #494E52;";
	insertLine = insertLine +  "\n background-color:#8D9091;";
	insertLine = insertLine +  "\n padding:8px;}";
	insertLine = insertLine +  "\n ul.tabs {margin:0px; padding:0px;}";
	insertLine = insertLine +  "\n ul.tabs li{";
	insertLine = insertLine +  "\n list-style:none;";
	insertLine = insertLine +  "\n display:inline;}";
	insertLine = insertLine +  "\n ul.tabs li a{";
	insertLine = insertLine +  "\n background-color:#4B4B4B;";
	insertLine = insertLine +  "\n color:#ffebb5;";
	insertLine = insertLine +  "\n padding:8px 14px 8px 14px;";
	insertLine = insertLine +  "\n text-decoration:none;";
	insertLine = insertLine +  "\n font-size:9px;";
	insertLine = insertLine +  "\n font-family:Verdana, Arial, Helvetica, sans-serif;";
	insertLine = insertLine +  "\n font-weight:bold;";
	insertLine = insertLine +  "\n text-transform:uppercase;";
	insertLine = insertLine +  "\n border:1px solid #464c54;}";
	insertLine = insertLine +  "\n ul.tabs li a:hover{";
	insertLine = insertLine +  "\n background-color:#2f343a;";
	insertLine = insertLine +  "\n border-color:#2f343a;}";
	insertLine = insertLine +  "\n ul.tabs li a.active{";
	insertLine = insertLine +  "\n background-color:#ffffff;";
	insertLine = insertLine +  "\n color:#282e32;";
	insertLine = insertLine +  "\n border:1px solid #464c54;";
	insertLine = insertLine +  "\n border-bottom: 1px solid #ffffff;}";
	insertLine = insertLine +  "\n .content{";
	insertLine = insertLine +  "\n background-color:#ffffff;";
	insertLine = insertLine +  "\n padding:10px;";
	insertLine = insertLine +  "\n border:1px solid #464c54;}";
	insertLine = insertLine +  "\n ul.tabs{";
	insertLine = insertLine +  "\n margin:0px; padding:0px;";
	insertLine = insertLine +  "\n margin-top:5px;";
	insertLine = insertLine +  "\n margin-bottom:6px;}";
	insertLine = insertLine +  "\n .content ul{";
	insertLine = insertLine +  "\n margin:0px;";
	insertLine = insertLine +  "\n padding:0px 20px 0px 20px;}";
	insertLine = insertLine +  "\n .content ul li{";
	insertLine = insertLine +  "\n list-style:none;";
	insertLine = insertLine +  "\n border-bottom:1px solid #d6dde0;";
	insertLine = insertLine +  "\n padding-top:15px;";
	insertLine = insertLine +  "\n padding-bottom:15px;";
	insertLine = insertLine +  "\n font-size:13px;}";
	insertLine = insertLine +  "\n .content ul li a{";
	insertLine = insertLine +  "\n text-decoration:none;";
	insertLine = insertLine +  "\n color:#3e4346;}";
	insertLine = insertLine +  "\n .content ul li a small{";
	insertLine = insertLine +  "\n color:#8b959c;";
	insertLine = insertLine +  "\n font-size:9px;";
	insertLine = insertLine +  "\n text-transform:uppercase;";
	insertLine = insertLine +  "\n font-family:Verdana, Arial, Helvetica, sans-serif;";
	insertLine = insertLine +  "\n position:relative;";
	insertLine = insertLine +  "\n left:4px;";
	insertLine = insertLine +  "\n top:0px;}";
	insertLine = insertLine +  "\n .content ul li:last-child {border-bottom:none;} \n";
	
	String strTemp = " ";
	for(int intI=1;intI<Driver.testModule.length;intI++){
		if(!Driver.testModule[intI].isEmpty()){
			strTemp = strTemp +",#content_"+ (intI+1);
		}
	}
	strTemp = strTemp +",#content_"+ (Driver.testModule.length + 1);
	insertLine = insertLine + strTemp.subSequence(2, strTemp.length());
	insertLine = insertLine +  "\n { display:none; }";
	strTemp=" ";
	
	insertLine = insertLine +  "\n </Style>";
	insertLine = insertLine +  "\n <script type=\"text/javascript\" src=\"https://www.google.com/jsapi\"></script>";
	insertLine = insertLine +  "\n <script type=\"text/javascript\">";	
	insertLine = insertLine +  "\n google.load(\"visualization\", \"1\", {packages:[\"corechart\"]});";
	insertLine = insertLine +  "\n google.setOnLoadCallback(fn_summary);";
	for(int intI=0;intI<Driver.testModule.length;intI++){
		if(!Driver.testModule[intI].isEmpty()){
			insertLine = insertLine +  "\n google.setOnLoadCallback(fn_module"+ (intI+1) +");";
		}
	}
	insertLine = insertLine +  "\n function fn_summary() {";
	insertLine = insertLine +  "\n var data = google.visualization.arrayToDataTable([";
	insertLine = insertLine +  "\n ['Module','TestCase Passed','TestCase Failed','Validation Passed','Validation Failed'],";
	for(int intI=0;intI<Driver.testModule.length;intI++){
		if(!Driver.testModule[intI].isEmpty()){
			strTemp = strTemp + "\n ['"+ Driver.testModule[intI].toString().trim() +"',parseInt(document.getElementById('testpass"+ (intI+1) +"').value),parseInt(document.getElementById('testfail"+ (intI+1) +"').value),parseInt(document.getElementById('validationpass"+ (intI+1) +"').value),parseInt(document.getElementById('validationfail"+ (intI+1) +"').value)],";
		}
	}
	insertLine = insertLine + strTemp.subSequence(1, (strTemp.length()-1));
	insertLine = insertLine +  "\n ]);";
	insertLine = insertLine +  "\n var options = {";
	insertLine = insertLine +  "\n title: 'iGen Framework Test Report',";
	insertLine = insertLine +  "\n hAxis: {title: 'Module', titleTextStyle: {color: 'red'}},";
	insertLine = insertLine +  "\n colors:['#69BE28','#E65032','#FFC300','#4B4B4B']";
	insertLine = insertLine +  "\n };";
	insertLine = insertLine +  "\n var chart = new google.visualization.ColumnChart(document.getElementById('chart_summary'));";
	insertLine = insertLine +  "\n chart.draw(data, options);";
	insertLine = insertLine +  "\n }";	
	for(int intI=0;intI<Driver.testModule.length;intI++){
		if(!Driver.testModule[intI].isEmpty()){
			insertLine = insertLine +  "\n function fn_module"+ (intI+1) +"() {";
			insertLine = insertLine +  "\n var data = google.visualization.arrayToDataTable([";
			insertLine = insertLine +  "\n ['Validation','Pass','Fail'],";
			insertLine = insertLine +  "\n ['Test Cases',parseInt(document.getElementById('testpass"+ (intI+1) +"').value),parseInt(document.getElementById('testfail"+ (intI+1) +"').value)],";
			insertLine = insertLine +  "\n ['Validations',parseInt(document.getElementById('validationpass"+ (intI+1) +"').value),parseInt(document.getElementById('validationfail"+ (intI+1) +"').value)]";
			insertLine = insertLine +  "\n ]);";
			insertLine = insertLine +  "\n var options = {";
			insertLine = insertLine +  "\n title: '"+ Driver.testModule[intI].trim() +"',";
			insertLine = insertLine +  "\n vAxis: {title: 'Validation',  titleTextStyle: {color: 'red'}},";
			insertLine = insertLine +  "\n colors:['#69BE28','#E65032']";
			insertLine = insertLine +  "\n };";
			insertLine = insertLine +  "\n var chart = new google.visualization.BarChart(document.getElementById('chart_module"+ (intI+1) +"'));";
			insertLine = insertLine +  "\n chart.draw(data, options);";
			insertLine = insertLine +  "\n }";
		}
	}	
	insertLine = insertLine +  "\n function tabSwitch(new_tab, new_content)";
	insertLine = insertLine +  "\n {	";
	for(int intI=0;intI<Driver.testModule.length;intI++){
		if(!Driver.testModule[intI].isEmpty()){
			insertLine = insertLine +  "\n document.getElementById('content_"+ (intI+1) +"').style.display = 'none';";
			insertLine = insertLine +  "\n document.getElementById('tab_"+ (intI+1) +"').className = '';";
		}
	}
	insertLine = insertLine +  "\n document.getElementById('content_"+ (Driver.testModule.length+1) +"').style.display = 'none';";
	insertLine = insertLine +  "\n document.getElementById('tab_"+ (Driver.testModule.length+1) +"').className = '';";
	insertLine = insertLine +  "\n document.getElementById(new_content).style.display = 'block';			";
	insertLine = insertLine +  "\n document.getElementById(new_tab).className = 'active';		";
	insertLine = insertLine +  "\n }";
	insertLine = insertLine +  "\n </script>";
	
	return insertLine;
	
}

/**
 * @param insertLine
 * @return
 */
private static String htmlResultTableCreate(String insertLine){
	
	insertLine = insertLine +  "\n </Table></li>";
	insertLine = insertLine +  "\n <li><TABLE borderColorLight=#008080 border=0 cellSpacing=1 cellPadding=3 width=\"100%\">";
	insertLine = insertLine +  "\n <TBody>";
	return insertLine;
	
}

/**
 * @param resultPath
 */
public static void htmlResultTableCreator(String resultPath){
	
	String insertLine="";
	byte[] contentInBytes;
	
	try (FileOutputStream oFso = new FileOutputStream(new File(resultPath),true)) {
		insertLine = insertLine +  "\n </TBody>";
		insertLine = htmlResultTableCreate(insertLine);
		contentInBytes = insertLine.getBytes();
		oFso.write(contentInBytes);
		oFso.flush();
		oFso.close();	
	} catch (IOException e) {
		e.printStackTrace();
	}

}

/**
 * @param strTName
 * @param strRName
 * @param strResName
 */
public static void insetResultTableRow(String[] strTName , String strRName, String resultPath){
	
	String insertLine="";
	byte[] contentInBytes;
	String line="";
	String input="";
	String strTemp="";
	
	if(strRName.equalsIgnoreCase("Header1")){
		insertLine = insertLine +  "\n </TBody>";
		insertLine = htmlResultTableCreate(insertLine);
	}
	insertLine = insertLine + insetResultTableRow(strTName,strRName);
	insertLine = insertLine +  "\n <!PerformLineInsertModule"+ (Driver.curModule+1) +">";	
	
	try (BufferedReader oFsoRead = new BufferedReader(new FileReader(resultPath));) {
		while ((line = oFsoRead.readLine()) != null){
			if(line.equalsIgnoreCase(" <!PerformLineInsertModule"+ (Driver.curModule+1) +">")){
				line = line.replace(" <!PerformLineInsertModule"+ (Driver.curModule+1) +">", insertLine);
			}else if(line.startsWith(" <!PerformTestCaseCountChanges>")){
				line = " <!PerformTestCaseCountChanges><script>";
				for(int intI=0;intI<Driver.testModule.length;intI++){
					strTemp = strTemp + "document.getElementById('testpass"+ (intI+1) +"').value="+ teststatus[intI][0] +";";
					strTemp = strTemp + "document.getElementById('testfail"+ (intI+1) +"').value="+ teststatus[intI][1] +";";
					strTemp = strTemp + "document.getElementById('validationpass"+ (intI+1) +"').value="+ teststatus[intI][2] +";";
					strTemp = strTemp + "document.getElementById('validationfail"+ (intI+1) +"').value="+ teststatus[intI][3] +";";
				}
				line = line + strTemp + "</script>";
			}
			input = input + line + '\n';
		}
		oFsoRead.close();
		
	}catch (IOException e) {
		e.printStackTrace();
	}
	
	try (FileOutputStream oFso = new FileOutputStream(new File(resultPath),false)) {
		contentInBytes = input.getBytes();
		oFso.write(contentInBytes);
		oFso.flush();
		oFso.close();	
	} catch (IOException e) {
		e.printStackTrace();
	}
	
}

/**
 * @param args
 * @param strType
 * @return
 */
private static String insetResultTableRow( String[] args , String strType){

	String insertLine="";
	String temp="";
	
	if((strType.contains("Header2"))|(strType.contains("Pass"))|(strType.contains("Fail"))){
		temp = " COLS = \"6\"";
	}
	
	insertLine = insertLine +  "\n <TR"+ temp +">";	
	for(int i=0;i<args.length;i++){
		switch(strType){
		
			case "Header":
				insertLine = insertLine +  "\n <TD vAlign=center align=middle width='3%' bgColor=#e1e1e1 height=30><FONT face=Verdana  color=#0033CC  size=1><B>"+ args[i].trim() +"</B></FONT></TD>";
				break;
				
			case "Header1":
				insertLine = insertLine +  "\n <TD vAlign=center align=left bgColor=#eeeeee width=\"30%\"><FONT size=2 face=CALIBRI><B>"+ args[i].trim() +"</B></FONT></TD>";
				insertLine = insertLine +  "\n <TD vAlign=center align=left bgColor=#eeeeee colSpan=5><FONT size=2 face=CALIBRI><B>"+ Driver.oModuleSheet.getCell(i,Driver.curTestRow).getContents().toString().trim() +"</B></FONT></TD></TR>";
				break;
				
			case "Header2":
				insertLine = insertLine +  "\n <TD vAlign=center align=center bgColor=#bbbbbb width=\"25%\"><FONT color=black size=2 face=CALIBRI><B>"+ args[i].trim() +"</B></FONT></TD>";
				break;
			
			case "Footer":
				insertLine = insertLine +  "\n <TD vAlign=center align=center width='3%' bgColor=#e1e1e1 height=30><FONT face=Verdana  color=#0033CC  size=2>"+ args[i].trim() +"</FONT></TD>";
				break;
				
			case "Pass": case "Passed":
				if((args[i].equalsIgnoreCase("Pass"))|(args[i].equalsIgnoreCase("Passed"))){
					insertLine = insertLine +  "\n <TD vAlign=center align=center bgColor=#eeeeee width=\"7%\"><FONT color=Green size=5 face=\"WINGDINGS 2\">P</FONT><FONT color=Green size=2 face=CALIBRI><B>"+ args[i].trim() +"</B></FONT></TD></TR>";
					teststatus[Driver.curModule][2] = teststatus[Driver.curModule][2]+1;
				}else{
					insertLine = insertLine +  "\n <TD vAlign=center align=left bgColor=#eeeeee width=\"30%\"><FONT size=2 face=CALIBRI> "+ args[i].trim() +"</FONT></TD>";
				}
				break;
			
			case "Fail": case "Failed":
				if((args[i].equalsIgnoreCase("Fail"))|(args[i].equalsIgnoreCase("Failed"))){
					insertLine = insertLine +  "\n <TD vAlign=center align=center bgColor=#eeeeee width=\"7%\"><FONT color=Red size=5 face=\"WINGDINGS 2\">O</FONT><FONT color=Red size=2 face=CALIBRI><B>"+ args[i].trim() +"</B></FONT></TD></TR>";
					teststatus[Driver.curModule][3]=teststatus[Driver.curModule][3]+1;
					Driver.teststatusFlag=false;
				}else{
					insertLine = insertLine +  "\n <TD vAlign=center align=left bgColor=#eeeeee width=\"30%\"><FONT size=2 face=CALIBRI> "+ args[i].trim() +"</FONT></TD>";
				}
				break;
				
		}
	}
	if(!strType.contains("Header1")){
		insertLine = insertLine +  "\n </TR>";
	}
	return insertLine;
}

}

##################################################################################################################
testcaseselector file
##################################################################################################################
package iGenFramework;

/*-----------------------------------------------------------------
Importing Support Classes
-----------------------------------------------------------------*/
import jxl.*;

import javax.swing.*;
import javax.swing.GroupLayout.*;
import javax.swing.LayoutStyle.*;

import java.awt.event.*;

@SuppressWarnings({ "unchecked", "rawtypes" })
public class TestCaseSelector {

	private JFrame frame;
	private JList lstTestContainer;
	private JList lstTestSelected;
	private static String[][] testcaseCollection;

	/**
	 * @param oBook
	 * @return
	 */
	public static String[][] ShowDialog(jxl.Workbook oBook){
		
		int testModuleCnt;
		TestCaseSelector jWindow;		
		
		testcaseCollection = new String[Driver.testModule.length][];
		for(testModuleCnt=0;testModuleCnt<Driver.testModule.length;testModuleCnt++){
			Sheet oSheet = oBook.getSheet(Driver.testModule[testModuleCnt]);
			try {
				jWindow = new TestCaseSelector(oSheet, testModuleCnt);
				jWindow.frame.setVisible(true);
				while(jWindow.frame.isVisible()){}
			} catch (Exception e) {
				System.out.println(e.toString());
			}
		}	
		return testcaseCollection;
	}

	/**
	 * @param oSheet
	 * @param oModule
	 */
	public TestCaseSelector(jxl.Sheet oSheet, int oModule) {
		initialize(oSheet, oModule);
	}

	/**
	 * @param oSheet
	 * @param oModule
	 */
	private void initialize(jxl.Sheet oSheet, final int oModule) {
		
		String testName;
		
		frame = new JFrame();
		frame.setBounds(100, 100, 501, 400);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setTitle(Driver.testModule[oModule]);
		
		JLabel lblDescription = new JLabel("Move Cases from Left Pane to Right to perform Test");
		lblDescription.setHorizontalAlignment(SwingConstants.CENTER);
		
		
		final DefaultListModel lstTestDefault = new DefaultListModel();
		final DefaultListModel lstTestNew = new DefaultListModel();
		
		for(int testRow=1;testRow<=oSheet.getRows()-1;testRow++){
			testName = Framework.fn_retrieveValue(oSheet, oSheet.getCell(0,0).getContents().toString().trim(),testRow).toString().trim();
			if(!testName.isEmpty()){
				lstTestDefault.addElement(testName);
			}			
		}
		
		lstTestContainer = new JList();		
		lstTestContainer.setModel(lstTestDefault);		
		lstTestSelected = new JList();
		lstTestSelected.setModel(lstTestNew);
		
		JButton btnMoveRight = new JButton("->");		
		btnMoveRight.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int index[] = lstTestContainer.getSelectedIndices();
				for(int i=0;i<index.length;i++){
					lstTestNew.addElement(lstTestContainer.getSelectedValue());
					lstTestDefault.removeElement(lstTestContainer.getSelectedValue());
				}
			}
		});
		
		JButton btnMoveAllRight = new JButton("-->");		
		btnMoveAllRight.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {				
				for(int i=0;i<lstTestContainer.getModel().getSize();i++){
					lstTestNew.addElement(lstTestContainer.getModel().getElementAt(i));
				}
				lstTestDefault.removeAllElements();
			}
		});
		JButton btnMoveLeft = new JButton("<-");		
		btnMoveLeft.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				int index[] = lstTestSelected.getSelectedIndices();
				for(int i=0;i<index.length;i++){
					lstTestDefault.addElement(lstTestSelected.getSelectedValue());
					lstTestNew.removeElement(lstTestSelected.getSelectedValue());
				}
			}
		});
		JButton btnMoveAllLeft = new JButton("<--");		
		btnMoveAllLeft.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				for(int i=0;i<lstTestSelected.getModel().getSize();i++){
					lstTestDefault.addElement(lstTestSelected.getModel().getElementAt(i));
				}
				lstTestNew.removeAllElements();
			}
		});
		
		JButton btnSubmit = new JButton("Submit");		
		btnSubmit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				testcaseCollection[oModule] = new String[lstTestSelected.getModel().getSize()];
				for(int i=0;i<lstTestSelected.getModel().getSize();i++){
					testcaseCollection[oModule][i] = lstTestSelected.getModel().getElementAt(i).toString();
				}
				frame.dispose();
			}
		});
		JButton btnCancel = new JButton("Cancel");		
		btnCancel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
			}
		});
		
		GroupLayout groupLayout = new GroupLayout(frame.getContentPane());
		groupLayout.setHorizontalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addContainerGap()
							.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING, false)
								.addGroup(groupLayout.createSequentialGroup()
									.addComponent(lstTestContainer, GroupLayout.PREFERRED_SIZE, 207, GroupLayout.PREFERRED_SIZE)
									.addPreferredGap(ComponentPlacement.UNRELATED)
									.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING)
										.addGroup(groupLayout.createSequentialGroup()
											.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING)
												.addComponent(btnMoveLeft, GroupLayout.DEFAULT_SIZE, 56, Short.MAX_VALUE)
												.addComponent(btnMoveAllRight, GroupLayout.DEFAULT_SIZE, 56, Short.MAX_VALUE)
												.addComponent(btnMoveRight, GroupLayout.DEFAULT_SIZE, 56, Short.MAX_VALUE))
											.addGap(5))
										.addGroup(groupLayout.createSequentialGroup()
											.addComponent(btnMoveAllLeft, GroupLayout.DEFAULT_SIZE, 55, Short.MAX_VALUE)
											.addPreferredGap(ComponentPlacement.RELATED)))
									.addComponent(lstTestSelected, GroupLayout.PREFERRED_SIZE, 189, GroupLayout.PREFERRED_SIZE))
								.addComponent(lblDescription, GroupLayout.PREFERRED_SIZE, 467, GroupLayout.PREFERRED_SIZE)))
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(135)
							.addComponent(btnSubmit, GroupLayout.PREFERRED_SIZE, 95, GroupLayout.PREFERRED_SIZE)
							.addGap(27)
							.addComponent(btnCancel, GroupLayout.PREFERRED_SIZE, 95, GroupLayout.PREFERRED_SIZE)))
					.addContainerGap())
		);
		groupLayout.setVerticalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addComponent(lblDescription, GroupLayout.PREFERRED_SIZE, 35, GroupLayout.PREFERRED_SIZE)
					.addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
						.addGroup(groupLayout.createSequentialGroup()
							.addPreferredGap(ComponentPlacement.RELATED)
							.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
								.addComponent(lstTestSelected, GroupLayout.PREFERRED_SIZE, 259, GroupLayout.PREFERRED_SIZE)
								.addComponent(lstTestContainer, GroupLayout.PREFERRED_SIZE, 259, GroupLayout.PREFERRED_SIZE)))
						.addGroup(groupLayout.createSequentialGroup()
							.addGap(45)
							.addComponent(btnMoveRight)
							.addGap(18)
							.addComponent(btnMoveAllRight)
							.addGap(56)
							.addComponent(btnMoveLeft)
							.addPreferredGap(ComponentPlacement.UNRELATED)
							.addComponent(btnMoveAllLeft)))
					.addGap(18)
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(btnSubmit)
						.addComponent(btnCancel))
					.addContainerGap(21, Short.MAX_VALUE))
		);
		frame.getContentPane().setLayout(groupLayout);
	}
}

##################################################################################################################
