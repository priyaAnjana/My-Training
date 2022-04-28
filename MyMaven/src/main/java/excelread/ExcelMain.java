package excelread;

import java.io.IOException;

public class ExcelMain {

	public static void main(String[] args) throws IOException {
		String s=ExcelCode.readStringData(0, 0);
		System.out.println(s);
		String s1=ExcelCode.integerData(1, 0);
		System.out.println(s1);
		
		// TODO Auto-generated method stub

	}

}
