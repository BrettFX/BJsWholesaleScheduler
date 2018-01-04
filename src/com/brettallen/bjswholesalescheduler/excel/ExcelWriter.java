package com.brettallen.bjswholesalescheduler.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import com.brettallen.bjswholesalescheduler.BJsWholesaleScheduler;
import jxl.CellType;
import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelWriter 
{
	private Workbook input;
	private WritableWorkbook copy;
    private File inFile;

	public WritableSheet sheet1;

	//Create a file number to append to each xlsFileName to be able to order files in file explorer
	private static int cashierFileNum = -1;
	private static int frontLineFileNum = 0;
	private static int overnightFileNum = 14;

    public ExcelWriter(InputStream path, BJsWholesaleScheduler.ScheduleTypes scheduleType, String day, String date) throws BiffException, IOException
	{		
		input = Workbook.getWorkbook(path);

		String scheduleName;

		int fileNum = 0;

		//Determine which file type the incoming request is and process accordingly (FrontLine even; Cashier odd
		if(scheduleType == BJsWholesaleScheduler.ScheduleTypes.FRONT_LINE) {
			if (frontLineFileNum >= 14)
				frontLineFileNum = 2;
			else
				frontLineFileNum +=2;

			fileNum = frontLineFileNum;
			scheduleName = "Front Line Schedule";
		}
		else if (scheduleType == BJsWholesaleScheduler.ScheduleTypes.CASHIER) {
			if(cashierFileNum >= 13)
				cashierFileNum = 1;
			else
				cashierFileNum += 2;

			fileNum = cashierFileNum;
			scheduleName = "Cashier Schedule";
		}
		else if (scheduleType == BJsWholesaleScheduler.ScheduleTypes.OVERNIGHT) {
			if(overnightFileNum >= 21)
				overnightFileNum = 15;
			else
				overnightFileNum++;

			fileNum = overnightFileNum;
			scheduleName = "Overnight Schedule";
		}
		else
			scheduleName = "Unknown Schedule";

        String xlsFileName = fileNum + "_" + scheduleName + " for " + day + ", " + date + ".xls";
		
		copy = Workbook.createWorkbook(new File(xlsFileName), input);		
		
		//Set the writable sheet to Sheet1 of the input excel file
		sheet1 = copy.getSheet(0);
	}
	
	public ExcelWriter(String path) throws BiffException, IOException
	{		
		input = Workbook.getWorkbook(new File(path));	
		
		inFile = new File("input.xls");
		
		copy = Workbook.createWorkbook(inFile, input);
		
		//Set the writable sheet to Sheet1 of the input excel file
		sheet1 = copy.getSheet(0);
	}
	
	public void deleteFirstColumn()
	{
		sheet1.removeColumn(0);
	}
	
	public void insertRow(int row)
	{
		sheet1.insertRow(row);
	}
	
	public boolean isEmptyCell(int col, int row)
	{
        WritableCell cell = sheet1.getWritableCell(col, row);

        return cell.getType() == CellType.EMPTY;
    }
	
	public void overwriteEmptyCell(int col, int row, String text, int cellFormat) throws WriteException
	{		
		//Set writable cell to the first name cell within FRONT LINE SUPERVISOR table		
		Label lbl = new Label(col , row, text, format(cellFormat));
		sheet1.addCell(lbl);
	}
	
	public File getXlsFile()
	{
		return inFile;
	}
	
	public String getXlsFilePath()
	{
		return inFile.getAbsolutePath();
	}
	
	private WritableCellFormat format(int choice) throws WriteException
	{
		WritableFont cellFont;
		WritableCellFormat cellFormat;
		
		switch(choice)
		{
		case 0://Day and Date
			//Set the font to Tahoma 12pt
			cellFont = new WritableFont(WritableFont.TAHOMA, 12);
			
			//Underline the text
			cellFont.setUnderlineStyle(UnderlineStyle.SINGLE);
			
			cellFormat = new WritableCellFormat(cellFont);
			
			return cellFormat;
		case 1://Names
			//Set the font to Tahoma 10pt
			cellFont = new WritableFont(WritableFont.TAHOMA, 10);
			
			//Set the cell format to have all boarders (prevents overwriting boarders)
			cellFormat = new WritableCellFormat(cellFont);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			
			return cellFormat;
		case 2: //Shift times
			//Set the font to Tahoma 9pt
			cellFont = new WritableFont(WritableFont.TAHOMA, 9);
			
			cellFormat = new WritableCellFormat(cellFont);
			cellFormat.setAlignment(Alignment.CENTRE);
			cellFormat.setBorder(Border.ALL, BorderLineStyle.THIN);
			
			return cellFormat;
		case 3: //Notes section for cashier schedule
			//Set the font to Tahoma 10pt
			cellFont = new WritableFont(WritableFont.TAHOMA, 12);
			
			//Underline the text
			cellFont.setUnderlineStyle(UnderlineStyle.SINGLE);
			
			cellFormat = new WritableCellFormat(cellFont);
			
			return cellFormat;
		default:
			return null;
		}
	}
	
	public void writeAndClose() throws IOException
	{
		copy.write();
		
		try 
		{
			copy.close();
		} 
		catch (WriteException e) 
		{
			e.printStackTrace();
		}
	}
}
