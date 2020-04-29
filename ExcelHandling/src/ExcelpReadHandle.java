import java.io.File;
import java.io.IOException;
import java.util.Scanner;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelpReadHandle 
{

	public void particularRowRangeData(int rowStartRange,int endRowRange,int colNum, String data) throws IOException, RowsExceededException, WriteException
	{
		
		File f=new File("C:\\\\Users\\\\dfd\\\\Desktop\\\\NishantTest2.xls");
		WritableWorkbook ww=Workbook.createWorkbook(f);
		WritableSheet ws=ww.createSheet("TestProg2",0);
		
		
		
	for(int i=rowStartRange-1;i<endRowRange;i++)
		
	{
		for(int j=0;j<colNum;j++)
			
		{
			
			
			Label la=new Label(j, i, data);
			ws.addCell(la);
		}
		}
		
	
		ww.write();
		ww.close();
	}
	
	
public static void main(String[] args) throws IOException, RowsExceededException, WriteException 
{

System.out.println("Enter Start Row for which you want to enter data");	
Scanner s=new Scanner(System.in);
int a=s.nextInt();
System.out.println("Enter End Row for which you want to print data");
int c=s.nextInt();
System.out.println("Enter Column Number in you want to enter data");
int b=s.nextInt();
System.out.println("Please enter the data");
String st=s.next();
ExcelpReadHandle ex=new ExcelpReadHandle();
ex.particularRowRangeData(a, c, b, st);

}
}
