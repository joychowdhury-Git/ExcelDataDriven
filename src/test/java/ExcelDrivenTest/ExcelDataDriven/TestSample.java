package ExcelDrivenTest.ExcelDataDriven;

import java.util.ArrayList;

public class TestSample {

	public static void main(String[] args) throws Exception {
		
		dataDrivenTest d= new dataDrivenTest();
		ArrayList<String> data = d.getdata("Add Profile");
		System.out.println(data.get(0));
		System.out.println(data.get(1));

		System.out.println(data.get(2));

		System.out.println(data.get(3));


	}

}
