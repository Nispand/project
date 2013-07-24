package format;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Vector;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.*;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import shapes.AdjacentBuildingData.*;

public class format
{

  
    public static void main (Vector arg1 , Vector arg2)
    {
        try
        {
        FileOutputStream fileout=new FileOutputStream("C:/Users/Nisu/Desktop/report2.xls");
        HSSFWorkbook workbook=new HSSFWorkbook();
        HSSFSheet worksheet=workbook.createSheet("Reportsheet");
        HSSFSheet sheet=workbook.createSheet("Adjacent Building Data");
        
        HSSFFont font=workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        HSSFCellStyle style=workbook.createCellStyle();
        style.setFont(font);
        HSSFCellStyle cstyle=workbook.createCellStyle();
        cstyle.setFillPattern(HSSFCellStyle.FINE_DOTS);
        cstyle.setFillBackgroundColor(new HSSFColor.DARK_YELLOW().getIndex());
        
        HSSFCellStyle rstyle=workbook.createCellStyle();
        rstyle.setFillPattern(HSSFCellStyle.FINE_DOTS);
        rstyle.setFillBackgroundColor(new HSSFColor.LIGHT_BLUE().getIndex());
		//Row 1
        HSSFRow row1=worksheet.createRow(0);
        HSSFCell cell1=row1.createCell(0);
        cell1.setCellValue("ANNEX-B");
        cell1.setCellStyle(style);
        worksheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 4));
        
		//Row2
        HSSFRow row2=worksheet.createRow(1);
        HSSFCell cell2=row2.createCell(0);
        cell2.setCellValue("FORMAT FOR CERTIFICATION OF BTS FOR COMPLIANCE OF THE EMF EXPOSURE LEVELS (CALCULATION OF EIRP/EIRPth)");
        
        cell2.setCellStyle(style);
        
        worksheet.addMergedRegion(new CellRangeAddress(1, 1, 0, 4));
        
		//Row 3
        HSSFRow row3=worksheet.createRow(2);
        HSSFCell cell3=row3.createCell(0);
        cell3.setCellValue("B-I : SITE DATA & TECHNICAL PARAMETERS");
        cell3.setCellStyle(style);
        worksheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 4));
		
		//Row 4
	HSSFRow row4=worksheet.createRow(3);
        HSSFCell cell4=row4.createCell(0);
        cell4.setCellValue("NAME OF BTS:");
        cell4.setCellStyle(style);
        worksheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 2));
		
	
        HSSFCell cell5=row4.createCell(3);
        cell5.setCellValue("KALOL_PRAYAGRAJ_797");
        cell5.setCellStyle(style);
		

        HSSFCell cell6=row4.createCell(4);
        cell6.setCellValue("O_GJ_GNG_KALOL_C_01_PRAYARAJ_APR");
        cell6.setCellStyle(style);
		
		//Row 5
        HSSFRow row5=worksheet.createRow(4);
        HSSFCell cell7=row5.createCell(0);
        cell7.setCellValue("B-I : SITE DATA & TECHNICAL PARAMETERS");
        
        cell7.setCellStyle(style);
        worksheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 2));
		

        HSSFCell cell8=row5.createCell(3);
        cell8.setCellValue(" ");

		

        HSSFCell cell9=row5.createCell(4);
        cell9.setCellValue(" ");

		
		//Row 6		
		HSSFRow row6=worksheet.createRow(5);
        HSSFCell cell10=row6.createCell(0);
        cell10.setCellValue(" ");

		

        HSSFCell cell11=row5.createCell(1);
        cell11.setCellValue(" ");

		

        HSSFCell cell12=row5.createCell(2);
        cell12.setCellValue(" ");

	

   //   HSSFCell cell13=row5.createCell(3);
   //     cell13.setCellValue("Operator1");

		

   //     HSSFCell cell14=row5.createCell(4);
   //     cell14.setCellValue("Operator2");

		

		HSSFRow row7=worksheet.createRow(6);
        HSSFCell cell15=row7.createCell(0);
        cell15.setCellValue("SITE DATA");
        cell15.setCellStyle(style);
      //  cell15.setCellStyle(vstyle);
        worksheet.addMergedRegion(new CellRangeAddress(6, 14, 0, 0));
		
		HSSFRow row16=worksheet.createRow(15);
        HSSFCell cell16=row16.createCell(0);
        cell16.setCellValue("TECHNICAL PARAMTERS");
        cell16.setCellStyle(style);
       // cell15.setCellStyle(vstyle);
        worksheet.addMergedRegion(new CellRangeAddress(15, 31, 0, 0));
		
		//Row 7
	
        HSSFCell cell17=row7.createCell(1);
        cell17.setCellValue("Item");
        cell17.setCellStyle(style);
        worksheet.addMergedRegion(new CellRangeAddress(6, 6, 1, 1));
		
	
        HSSFCell cell18=row7.createCell(2);
        cell18.setCellValue("Units");
        cell18.setCellStyle(style);
		
		//Column 2 (B)
		
		HSSFRow row8=worksheet.createRow(7);
        HSSFCell cell19=row8.createCell(1);
        cell19.setCellValue("Site ID");
        
		
		HSSFRow row9=worksheet.createRow(8);
		HSSFCell cell20=row9.createCell(1);
		cell20.setCellValue("Name");
		worksheet.addMergedRegion(new CellRangeAddress(8, 8, 1, 1));
		
		HSSFRow row10=worksheet.createRow(9);
		HSSFCell cell21=row10.createCell(1);
		cell21.setCellValue("Date of Commissioning");
		worksheet.addMergedRegion(new CellRangeAddress(9, 9, 1, 1));
		
		HSSFRow row11=worksheet.createRow(10);
		HSSFCell cell22=row11.createCell(1);
		cell22.setCellValue("Address");
	
		
		HSSFRow row12=worksheet.createRow(11);
		HSSFCell cell23=row12.createCell(1);
		cell23.setCellValue("Lat / Long");
		worksheet.addMergedRegion(new CellRangeAddress(11, 11, 1, 1));
		
		HSSFRow row13=worksheet.createRow(12);
		HSSFCell cell24=row13.createCell(1);
		cell24.setCellValue("RTT / GBT");
	
		
		HSSFRow row14=worksheet.createRow(13);
		HSSFCell cell25=row14.createCell(1);
		cell25.setCellValue("Building Height AGL");
	
		
		HSSFRow row15=worksheet.createRow(14);
		HSSFCell cell26=row15.createCell(1);
		cell26.setCellValue("Antenna Height AGL");
	
		
	
		HSSFCell cell27=row16.createCell(1);
		cell27.setCellValue("System Type (GSM/CDMA/UTMS)");
	
		
		HSSFRow row17=worksheet.createRow(16);
		HSSFCell cell28=row17.createCell(1);
		cell28.setCellValue("Base Channel Frequency");
	
		
		HSSFRow row18=worksheet.createRow(17);
		HSSFCell cell29=row18.createCell(1);
		cell29.setCellValue("Carriers / Sector (Worst)");
	
		
		HSSFRow row19=worksheet.createRow(18);
		HSSFCell cell30=row19.createCell(1);
		cell30.setCellValue("Make and Model of Antenna");
	
		
		HSSFRow row20=worksheet.createRow(19);
		HSSFCell cell31=row20.createCell(1);
		cell31.setCellValue("Antenna Gain");
	
		
		HSSFRow row21=worksheet.createRow(20);
		HSSFCell cell32=row21.createCell(1);
		cell32.setCellValue("Total Tilt");
	
		
		HSSFRow row22=worksheet.createRow(21);
		HSSFCell cell33=row22.createCell(1);
		cell33.setCellValue("Vertical Beamwidth");
	
		
		HSSFRow row23=worksheet.createRow(22);
		HSSFCell cell34=row23.createCell(1);
		cell34.setCellValue("Side Lobe Attenuation");
	
		
		HSSFRow row24=worksheet.createRow(23);
		HSSFCell cell35=row24.createCell(1);
		cell35.setCellValue("Tx Power");
		worksheet.addMergedRegion(new CellRangeAddress(23, 23, 1, 1));
		
		HSSFRow row25=worksheet.createRow(24);
		HSSFCell cell36=row25.createCell(1);
		cell36.setCellValue("Combiner Loss");
	
		
		HSSFRow row26=worksheet.createRow(25);
		HSSFCell cell37=row26.createCell(1);
		cell37.setCellValue("RF Cable Length");
	
		
		HSSFRow row27=worksheet.createRow(26);
		HSSFCell cell38=row27.createCell(1);
		cell38.setCellValue("Unit Loss");
	
		
		HSSFRow row28=worksheet.createRow(27);
		HSSFCell cell39=row28.createCell(1);
		cell39.setCellValue("EIRP (Base Channel)");
	
		
		HSSFRow row29=worksheet.createRow(28);
		HSSFCell cell40=row29.createCell(1);
		cell40.setCellValue("DTX factor");
	
		
		HSSFRow row30=worksheet.createRow(29);
		HSSFCell cell41=row30.createCell(1);
		cell41.setCellValue("ATPC factor");
	
		
		HSSFRow row31=worksheet.createRow(30);
		HSSFCell cell42=row31.createCell(1);
		cell42.setCellValue("EIRC (TCH) incl DTX, ATPC");
	
		
		HSSFRow row32=worksheet.createRow(31);
		HSSFCell cell43=row32.createCell(1);
		cell43.setCellValue("EIRP (Total)");
	
		
		//Column 3 (C)
		
	
		HSSFCell cell44=row8.createCell(2);
                cell44.setCellValue(" ");
        
		
	
		HSSFCell cell45=row9.createCell(2);
		cell45.setCellValue(" ");
	
		
	
		HSSFCell cell46=row10.createCell(2);
		cell46.setCellValue(" ");
	
		
	
		HSSFCell cell47=row11.createCell(2);
		cell47.setCellValue(" ");
	
		
	
		HSSFCell cell48=row12.createCell(2);
		cell48.setCellValue(" ");
	
		
	
		HSSFCell cell49=row13.createCell(2);
		cell49.setCellValue(" ");
	
		
	
		HSSFCell cell50=row14.createCell(2);
		cell50.setCellValue("(m)");
	
		
	
		HSSFCell cell51=row15.createCell(2);
		cell51.setCellValue("(m)");
	
		

		HSSFCell cell52=row16.createCell(2);
		cell52.setCellValue(" ");

		

		HSSFCell cell53=row17.createCell(2);
		cell53.setCellValue("(MHz)");

		

		HSSFCell cell54=row18.createCell(2);
		cell54.setCellValue(" ");

		

		HSSFCell cell55=row19.createCell(2);
		cell55.setCellValue(" ");

		

		HSSFCell cell56=row20.createCell(2);
		cell56.setCellValue("(dBi)");

		

		HSSFCell cell57=row21.createCell(2);
		cell57.setCellValue("(Deg)");

		

		HSSFCell cell58=row22.createCell(2);
		cell58.setCellValue("(Deg)");

		

		HSSFCell cell59=row23.createCell(2);
		cell59.setCellValue("(db)");

		

		HSSFCell cell60=row24.createCell(2);
		cell60.setCellValue("(dBm)");

		

		HSSFCell cell61=row25.createCell(2);
		cell61.setCellValue("(db)");

		

		HSSFCell cell62=row26.createCell(2);
		cell62.setCellValue("(m)");

		

		HSSFCell cell63=row27.createCell(2);
		cell63.setCellValue("(dB/100m)");

		

		HSSFCell cell64=row28.createCell(2);
		cell64.setCellValue("(W)");

		

		HSSFCell cell65=row12.createCell(2);
		cell65.setCellValue(" ");

		

		HSSFCell cell66=row30.createCell(2);
		cell66.setCellValue(" ");

		

		HSSFCell cell67=row31.createCell(2);
		cell67.setCellValue("(W)");

		

		HSSFCell cell68=row12.createCell(2);
		cell68.setCellValue("(W)");

		
		//Column 4 (D)
		
 	int i=3,j=0,op=1;
        int loop=arg1.size();
	while(j<loop)
	{
		
	      HSSFCell cell13=row6.createCell(i);
		        cell13.setCellValue("Operator "+ op);
                        cell13.setCellStyle(style);
        //      Operator
		HSSFCell cell69=row7.createCell(i);
                       
		        cell69.setCellValue((String)arg1.elementAt(j++));
                         cell69.setCellStyle(cstyle);
                        //cell69.setCellStyle(style);
        //      Site ID
		HSSFCell cell70=row8.createCell(i);
                           cell70.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
		       
        //      Name
                HSSFCell cell71=row9.createCell(i);
		cell71.setCellValue((String)arg1.elementAt(j++));
	
		
	//	Date_Of_Commisioning (Need Correction-not perfect)
		HSSFCell cell72=row10.createCell(i);
		cell72.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Address
		HSSFCell cell73=row11.createCell(i);
		cell73.setCellValue((String)arg1.elementAt(j++));
		worksheet.addMergedRegion(new CellRangeAddress(10, 10, 3, 4));
		
	//	Latitude/Longitude
		HSSFCell cell74=row12.createCell(i);
		cell74.setCellValue((String)arg1.elementAt(j++));
		worksheet.addMergedRegion(new CellRangeAddress(11, 11, 3, 4));
		
	//	RTT/GBT
		HSSFCell cell75=row13.createCell(i);
		cell75.setCellValue((String)arg1.elementAt(j++));

		
	//	Building Height
		HSSFCell cell76=row14.createCell(i);
		cell76.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());

		
	//	Antenna Height
		HSSFCell cell77=row15.createCell(i);
		cell77.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());

		
	//	System Type
		HSSFCell cell78=row16.createCell(i);
		cell78.setCellValue((String)arg1.elementAt(j++));
		
	//	Base Channel Frequency(bcf)
		HSSFCell cell79=row17.createCell(i);
		cell79.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());

		
	//	Carrier (car)/Sector(Worst)
		HSSFCell cell80=row18.createCell(i);
		cell80.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
		
	//	Make and Model of Antenna
		HSSFCell cell81=row19.createCell(i);
		cell81.setCellValue((String)arg1.elementAt(j++));

		
	//	Antenna Gain
		HSSFCell cell82=row20.createCell(i);
		cell82.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Total Tilt
		HSSFCell cell83=row21.createCell(i);
		cell83.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Vertical bandwidth
		HSSFCell cell84=row22.createCell(i);
		cell84.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Side Loab Attenuation
		HSSFCell cell85=row23.createCell(i);
		cell85.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	TX_Power
		HSSFCell cell86=row24.createCell(i);
		cell86.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Combiner Loss
		HSSFCell cell87=row25.createCell(i);
		cell87.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Cable Lenght
		HSSFCell cell88=row26.createCell(i);
		cell88.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
		
	//	Unit_Loss
		HSSFCell cell89=row27.createCell(i);
		cell89.setCellValue(((Double)arg1.elementAt(j++)).doubleValue());
	
                i++;
                op++;
}		
        int size2=arg2.size(),fin=0;
        int k=3;
        while(fin<size2)
        {
        //	EIRP
		HSSFCell cell90=row28.createCell(k);
		cell90.setCellValue(((Double)arg2.elementAt(fin++)).doubleValue());
                cell90.setCellStyle(rstyle);
		
	//	DTX_Factor
		HSSFCell cell91=row29.createCell(k);
		cell91.setCellValue(((Double)arg2.elementAt(fin++)).doubleValue());
                cell91.setCellStyle(rstyle);
		
	//	ATPC_Factor
		HSSFCell cell92=row30.createCell(k);
		cell92.setCellValue(((Double)arg2.elementAt(fin++)).doubleValue());
                cell92.setCellStyle(rstyle);
		
	//	EIRP_TCH
		HSSFCell cell93=row31.createCell(k);
		cell93.setCellValue(((Double)arg2.elementAt(fin++)).doubleValue());
		cell93.setCellStyle(rstyle);
	//	EIRP_TOTAL
		HSSFCell cell94=row32.createCell(k);
		cell94.setCellValue(((Double)arg2.elementAt(fin++)).doubleValue());  
                cell94.setCellStyle(rstyle);
                k++;
        }
        for(int l=0;l<j;l++)
        {
            worksheet.autoSizeColumn((short)l);
        }
	
        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
        HSSFSimpleShape a = patriarch.createSimpleShape( new HSSFClientAnchor(100, 100, 1000, 200, (short)0, 0, (short)0, 0) );
        //HSSFSimpleShape shape1 = patriarch.createSimpleShape(a);
        a.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        a.setNoFill(true);
        System.out.println("Cirlce drawn");
        
        
     
             
	workbook.write(fileout);
        }
        catch(FileNotFoundException fn)
        {
            System.out.println("File Not Found"+fn);
        }
        catch(IOException io)
        {
            System.out.println("IO Excepttion"+io);
        }
        
    }
}