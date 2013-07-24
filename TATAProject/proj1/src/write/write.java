package write;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import javax.swing.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.hssf.util.HSSFColor;



public class write
{
    public static void main(String args[])
    {
        try
        {
        FileOutputStream fileout=new FileOutputStream("C:/Users/Nisu/Desktop/report1.xls");
        HSSFWorkbook workbook=new HSSFWorkbook();
        HSSFSheet worksheet=workbook.createSheet("Reportsheet");
       // Formulacoding f=new Formulacoding();
        
        //index from 0,0
        HSSFRow row0=worksheet.createRow(16);
        HSSFCell cell0=row0.createCell(3);
        JOptionPane.showMessageDialog(null,"writing bcf::"+args[0]);
        cell0.setCellValue(args[0]);
       // worksheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
       // HSSFCellStyle cellstyle=workbook.createCellStyle();
        
        HSSFRow row1=worksheet.createRow((short)17);
        HSSFCell cell1=row1.createCell(3);
        cell1.setCellValue(args[1]);
        //worksheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
      
        HSSFRow row2=worksheet.createRow(19);
        HSSFCell cell2=row2.createCell(3);
        cell2.setCellValue(args[2]);
        //worksheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 3));
        
        HSSFRow row3=worksheet.createRow(20);
        HSSFCell cell3=row3.createCell(3);
        cell3.setCellValue(args[3]);
       // worksheet.addMergedRegion(new CellRangeAddress(2, 2,0, 2));
        
       // HSSFCell cell4=row3.createCell();
        //cell4.setCellValue("KALOL_PRAYAGRAJ_797");
        
        //HSSFCell cell5=row3.createCell(4);
        //cell5.setCellValue("O_GJ_GNG_KALOL_C_01_PRAYARAJ_APR");
        
        HSSFRow row4=worksheet.createRow(21);
        HSSFCell cell6=row4.createCell(3);
        cell6.setCellValue(args[4]);
        //worksheet.addMergedRegion(new CellRangeAddress(5, 5,0, 2));
        
        
        
        HSSFRow row5=worksheet.createRow(22);
        HSSFCell cell7=row5.createCell(3);
        cell7.setCellValue(args[5]);
        //HSSFCell cell8=row5.createCell(4);
      //  cell8.setCellValue("OPERATOR 2");
        
        HSSFRow row6=worksheet.createRow(23);
        HSSFCell cell9=row6.createCell(3);
        //worksheet.addMergedRegion(new CellRangeAddress(6, 14,0, 0));
        cell9.setCellValue(args[6]);
        //HSSFCell cell10=row6.createCell(1);
        //cell10.setCellValue("item");
       // HSSFCell cell11=row6.createCell(2);
        //cell11.setCellValue("units");
        //HSSFCell cell12=row6.createCell(3);
        //cell12.setCellValue("TATA CDMA");
//        cellstyle.setFillBackgroundColor((short)2);
        //HSSFCell cell13=row6.createCell(4);
       // cell13.setCellValue("TATA GSM");
        HSSFRow row7=worksheet.createRow(24);
        HSSFCell cell10=row7.createCell(3);
        cell10.setCellValue(args[7]);
        
        HSSFRow row8=worksheet.createRow(25);
        HSSFCell cell11=row8.createCell(3);
        cell11.setCellValue(args[8]);
        
        HSSFRow row9=worksheet.createRow(26);
        HSSFCell cell12=row9.createCell(3);
        cell12.setCellValue(args[9]);
      
        //HSSFRow row7=worksheet.createRow(27);
        
        
        
        workbook.write(fileout);

        }
        catch(FileNotFoundException fn)
        {
            System.out.println("File not found"+fn);
        }
        catch(IOException io)
        {
            System.out.println("IO exception");
        }
    }
}