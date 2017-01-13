package bhead_report;

// This program is called from UNIX script op_gn_cmp_report_bhead.sh
import java.io.*;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Gn_cmp_report_bhead_main 
{
   public static void main(String[] args) throws Exception 
   {
      // Output file name
      String inputFileName = args[0];
      String outputFileName = args[1];
      String strLine;
      int i=1;
      String mainItemUdac;
      String mainItemUdacNoSpaces;
      String headingCode;
      String headingName;
      final String BHED_UDAC = "BHED";

      // Open input file
      FileInputStream fstream = new FileInputStream(inputFileName);
      BufferedReader br = new BufferedReader(new InputStreamReader(fstream));

      // Create blank workbook
      XSSFWorkbook workbook = new XSSFWorkbook(); 
      // Create a blank sheet
      XSSFSheet spreadsheet = workbook.createSheet(" BHED Headings "); 
      // Create row object
      XSSFRow row;
      // This data needs to be written (Object[])
      Map < String, Object[] > headingInfo = new TreeMap < String, Object[] >(); 
      headingInfo.put( "1", new Object[] {"HEADING CODE", "HEADING NAME"}); 
      
      //Read File Line By Line
      while ((strLine = br.readLine()) != null)   {
          mainItemUdac = strLine.substring(23, 33);
          mainItemUdacNoSpaces = mainItemUdac.replaceFirst("\\s+$", "");
          headingCode = strLine.substring(44, 52);
          headingName = strLine.substring(1072);
          if ( mainItemUdacNoSpaces.equals (BHED_UDAC) ) {
              i++;
              headingInfo.put(Integer.toString(i) , new Object[] { headingCode, headingName });
          }
      }
      //Close the input stream
      br.close();


      //Iterate over data and write to sheet
      Set < String > keyid = headingInfo.keySet();
      int rowid = 0;
      for (String key : keyid)
      {
         row = spreadsheet.createRow(rowid++);
         Object [] objectArr = headingInfo.get(key);
         int cellid = 0;
         for (Object obj : objectArr)
         {
            Cell cell = row.createCell(cellid++);
            cell.setCellValue((String)obj);
         }
      }
      spreadsheet.autoSizeColumn(0);
      spreadsheet.autoSizeColumn(1);

      //Write the workbook in file system
      FileOutputStream out = new FileOutputStream(new File(outputFileName)); 
      workbook.write(out);
      out.close();
      System.out.println(outputFileName + " written successfully" ); 
      System.exit(0);
   }
}

