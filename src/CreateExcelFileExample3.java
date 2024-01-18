    import  java.io.*;  
    import  org.apache.poi.hssf.usermodel.HSSFSheet;  
    import  org.apache.poi.hssf.usermodel.HSSFWorkbook;  
    import  org.apache.poi.hssf.usermodel.HSSFRow;  
    public class CreateExcelFileExample3  
    {  
    public static void main(String[]args)   
    {  
    try   
    {  
     
    String fnm = "./FilesData/WriteHere1.xls";  
    // HSSFWorkbook  
    HSSFWorkbook workbook = new HSSFWorkbook();  
    //creatSheet()   
    HSSFSheet sheet = workbook.createSheet("TestSheet1");   
    //createRow() method  
    HSSFRow rowhead = sheet.createRow((short)0);  
    //creating cell by using the createCell() method and setting the values to the cell by using the setCellValue() method  
    rowhead.createCell(0).setCellValue("No");  
    rowhead.createCell(1).setCellValue("Name");  
    rowhead.createCell(2).setCellValue("Accnumber");  
    rowhead.createCell(3).setCellValue("email");  
    rowhead.createCell(4).setCellValue("Amount");  
    //creating the 1st row  
    HSSFRow row = sheet.createRow((short)1);  
    //inserting data in the first row  
    row.createCell(0).setCellValue("101");  
    row.createCell(1).setCellValue("Purabi Mishra");  
    row.createCell(2).setCellValue("111111");  
    row.createCell(3).setCellValue("p@gmail.com");  
    row.createCell(4).setCellValue("900000.00");  
    //creating the 2nd row  
    HSSFRow row1 = sheet.createRow((short)2);  
    //inserting data in the second row  
    row1.createCell(0).setCellValue("1002");  
    row1.createCell(1).setCellValue("Gayatri Mishra");  
    row1.createCell(2).setCellValue("1134");  
    row1.createCell(3).setCellValue("g@gmail.com");  
    row1.createCell(4).setCellValue("45000.00");  
    
    HSSFRow row2 = sheet.createRow((short)3);
    row2.createCell(0).setCellValue("10099");  
    row2.createCell(1).setCellValue("Mishra");  
    row2.createCell(2).setCellValue("7865");  
    row2.createCell(3).setCellValue("p@gmail.com");  
    row2.createCell(4).setCellValue("897.90");  
    
    HSSFRow row3 = sheet.createRow((short)4);
    row3.createCell(0).setCellValue("5645");  
    row3.createCell(1).setCellValue("test");  
    row3.createCell(2).setCellValue("675");  
    row3.createCell(3).setCellValue("ytiutuotg@gmail.com");  
    row3.createCell(4).setCellValue("89097.90");  
  
    
    
    
    FileOutputStream fileOut = new FileOutputStream(fnm);  
    workbook.write(fileOut);  
    //closing the Stream  
    fileOut.close();  
    //closing the workbook  
    workbook.close();  
    //prints the message on the console  
    System.out.println("Excel file has been generated successfully.");  
    }   
    catch (Exception e)   
    {  
    e.printStackTrace();  
    }  
    }
    }
   