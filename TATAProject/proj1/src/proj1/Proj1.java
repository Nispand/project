/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package proj1;
/**
 *
 * @author Nisu
 */
import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Array;
import java.util.Iterator;
import java.util.Vector;
import jxl.DateCell;
import java.util.Date;
import org.apache.poi.hssf.model.Sheet;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;



public class Proj1 
{

    /**
     * @param args the command line arguments
     */
    public static void main(Vector arg) 
    {
        try
        {       
            
               
            InputStream input=new BufferedInputStream(new FileInputStream("C:/Users/Nisu/Desktop/test1.xls"));
        POIFSFileSystem fs=new POIFSFileSystem(input);
        HSSFWorkbook wb=new HSSFWorkbook(fs);
        HSSFSheet sheet=wb.getSheetAt(0);
        int i=0;
        Iterator rows=sheet.rowIterator();
        HSSFRow row1=(HSSFRow)rows.next();
        while(rows.hasNext())
        {
            HSSFRow row=(HSSFRow)rows.next();
            System.out.println("\n");
            Iterator cells=row.cellIterator();
            while(cells.hasNext())
            {
                HSSFCell cell=(HSSFCell)cells.next();
            /*    if(HSSFDateUtil.isCellDateFormatted(cell))
                {
                    Date d=new Date();
                    d=cell.getDateCellValue();
                    arg.addElement(d);
                }
*/          
                if(HSSFCell.CELL_TYPE_NUMERIC==cell.getCellType())
                {
                    System.out.println(cell.getNumericCellValue() + "  ");
                    Double x=new Double(cell.getNumericCellValue());
                    arg.addElement(x);
                    i++;
                    
                }
                else if(HSSFCell.CELL_TYPE_STRING==cell.getCellType())
                {
                    System.out.print(cell.getStringCellValue()+ "  ");
                    String s=new String(cell.getStringCellValue());
                    arg.addElement(s);
                    i++;
                }
           /*     else if(HSSFCell.CELL_TYPE_BOOLEAN==cell.getCellType())
                {
                    System.out.println(cell.getBooleanCellValue()+"  ");
                }
            */
                else
    if(HSSFCell.CELL_TYPE_BLANK==cell.getCellType())
                    {
                        System.out.println("BLANK    ");
                    }
                    else
                    {
                        System.out.println("Unknown Cell Type");
                    }
            }
        }
        }
       catch(FileNotFoundException fn)
        {
            System.out.println("File not found"+fn);
        }
       
        catch(IOException io)
        {
            System.out.println("IOException "+ io);
            io.printStackTrace();
        }
        
        
    }
                
}
