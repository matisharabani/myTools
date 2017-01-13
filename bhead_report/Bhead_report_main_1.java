package bhead_report;

import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.Date;
//import java.util.Calendar;
//import java.text.DateFormat;
//import java.text.SimpleDateFormat;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Bhead_report_main_1 {
	static Date today = new Date();
	
	public static void main(String[] args) throws Exception {
		// Display current date
		System.out.println("Today is: " + today);
		
		// Create file object for the directory
		File dir=new File("C:\\DATAFILES\\CM");
		File[] files = dir.listFiles(new FileFilter() {
		    public boolean accept(File file) {
		    	// Filter *.CMP files modified within the last 100 days.
		    	long fileAge=(System.currentTimeMillis()- file.lastModified())/1000/60/60/24;
		        return file.getName().endsWith(".CMP") && fileAge < 100;
		    }
		});   
		if (files != null) {
			System.out.println("Number of files in dir: " + files.length);
	        for (File file : files) { 
	        	handleFile(file);	
	        	
	        }
		}
	   		
	   //DateFormat dateFormat = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");
	   //SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm:ss");	   
	   //System.out.println("SimpleDateFormat: " + sdf.format(today)); 
	   //System.out.println("Today in milliseconds: " + System.currentTimeMillis());
	   	   
   }
   
 
   public static void handleFile(File file) {
	   
	   String inputFileName = file.getPath();
	   String outputFileName = inputFileName + ".xlsx";
	   String oldOutputFileName = file.getParent() + "\\DEX_" + file.getName() + ".xlsx";
	   String strLine;
	   int i=1;
	   String mainItemUdac;
	   String mainItemUdacNoSpaces;
	   String headingCode;
	   String headingName;
	   final String BHED_UDAC = "BHED";
	   System.out.println("");
		  
	   System.out.println("Input File Name: " + inputFileName);
	   System.out.println("Output File Name: " + outputFileName);
	   System.out.println("Old Output File Name: " + oldOutputFileName);
		
	   System.out.println("Renaming " + outputFileName + " to " + oldOutputFileName);
	   File outputFile = new File(outputFileName);
	   outputFile.renameTo(new File (oldOutputFileName));
     	  
     	  
	   // Open input file
	   try {
		   FileInputStream fstream = new FileInputStream(inputFileName);
		   BufferedReader br = new BufferedReader(new InputStreamReader(fstream));
  		  
		   // This data needs to be written (Object[])
		   Map < String, Object[] > headingInfo = new TreeMap < String, Object[] >(); 
  		  
		   headingInfo.put( "1", new Object[] {"HEADING CODE", "HEADING NAME"}); 

		   //Read File Line By Line
		   try {
			   while ((strLine = br.readLine()) != null) {				   
				   mainItemUdac = strLine.substring(23, 33);
				   // Remove trailing spaces
				   mainItemUdacNoSpaces = mainItemUdac.replaceFirst("\\s+$", "");
				   headingCode = strLine.substring(44, 52);
				   headingName = strLine.substring(1072);
				   if ( mainItemUdacNoSpaces.equals (BHED_UDAC) ) {
					   i++;
					   headingInfo.put(Integer.toString(i) , new Object[] { headingCode, headingName });
					   //System.out.println("Heading: " + headingCode + "-" + headingName);
				   }
			   }
			   writeToXL(headingInfo, outputFileName);
		   } 
		   catch (Exception e) {
			   System.out.println ("Error reading from " + inputFileName);
			   e.printStackTrace();
		   }
		   br.close();
	   } 
	   catch (FileNotFoundException e) {
		   System.out.println ("File not found " + inputFileName);
		   e.printStackTrace();
	   } 
	   catch (IOException e) {
		   System.out.println ("Error opening " + inputFileName);
		   e.printStackTrace();
	   }	   
   }
   
   public static void writeToXL (Map < String, Object[] > headingInfo, String outputFileName) {
       // Create blank workbook
       XSSFWorkbook workbook = new XSSFWorkbook(); 
       // Create a blank sheet
       XSSFSheet spreadsheet = workbook.createSheet(" BHED Headings "); 
       // Create row object
       XSSFRow row;
	   
       //Iterate over data and write to sheet
       Set < String > keyid = headingInfo.keySet();
       int rowid = 0;
       for (String key : keyid) {
           row = spreadsheet.createRow(rowid++);
           Object [] objectArr = headingInfo.get(key);
           int cellid = 0;
           for (Object obj : objectArr) {
        	   //System.out.println("Heading Object: " + (String)obj + "-" + obj.toString());
               Cell cell = row.createCell(cellid++);
               cell.setCellValue((String)obj);
           }
        }
        spreadsheet.autoSizeColumn(0);
        spreadsheet.autoSizeColumn(1);

       //Write the workbook in file system
        try {
			FileOutputStream out = new FileOutputStream(new File(outputFileName)); 
			try {               
			    workbook.write(out);
			    System.out.println(outputFileName + " written successfully" ); 
			}
			catch (IOException e) {
			    System.out.println("Error writing " + outputFileName);
			    System.exit(1); 
			}
			finally {
			  out.close();
			}
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}  
   }
   
}


