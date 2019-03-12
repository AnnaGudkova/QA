import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

import com.lowagie.text.Cell;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Font;
import com.lowagie.text.Paragraph;
import com.lowagie.text.pdf.PdfWriter;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.pdf.PdfPTable;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.text.ParseException; 	
import java.time.LocalDate;
import java.time.Period;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Random;

public class Data {
    public static void main(String[] args) throws ParseException, DocumentException, IOException {   
        // �������� Excel ����� � ������
        HSSFWorkbook workbook = new HSSFWorkbook();
        // �������� ����� 
        HSSFSheet sheet = workbook.createSheet("������ �������������");
        // ���������� ������ �������
        List<String> dataList = fillData();
        // ������� ��� �����
        int rowNum = 0;
        // ����� �������
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue("���");
        row.createCell(1).setCellValue("�������");
        row.createCell(2).setCellValue("��������");
        row.createCell(3).setCellValue("�������");
        row.createCell(4).setCellValue("���");
        row.createCell(5).setCellValue("���� ��������");
        row.createCell(6).setCellValue("���");
        row.createCell(7).setCellValue("�������� ������");
        row.createCell(8).setCellValue("������");
        row.createCell(9).setCellValue("�������");
        row.createCell(10).setCellValue("�����");
        row.createCell(11).setCellValue("�����");
        row.createCell(12).setCellValue("���");
        row.createCell(13).setCellValue("��������");
        // �������� pdf �����
        Document document = new Document(); 
        PdfWriter writer;
        BaseFont font = null;
        // ����� ��� ����������� ������� ����
        font = BaseFont.createFont("arial.ttf","cp1251",BaseFont.EMBEDDED);		
        Font myFont = new Font(font,8);
        File pdfFile = new File("Data.pdf");
		String pathPdfFile = pdfFile.getAbsolutePath();
		writer = PdfWriter.getInstance(document, new FileOutputStream(pdfFile));
		document.open();
		PdfPTable table=new PdfPTable(14);
		// ������ ������� �� ���� ����
        table.setWidthPercentage(100);
        // ����� �������
        List<String> headings = new ArrayList<String>();
        headings.add(0, "���");
        headings.add(1, "�������");
        headings.add(2, "��������");
        headings.add(3, "�������");
        headings.add(4, "���");
        headings.add(5, "���� ��������");
        headings.add(6, "���");
        headings.add(7, "�������� ������");
        headings.add(8, "������");
        headings.add(9, "�������");
        headings.add(10, "�����");
        headings.add(11, "�����");
        headings.add(12, "���");
        headings.add(13, "��������");
        for (int i=0; i<=13; i++) {
        	table.addCell(new Paragraph(headings.get(i),myFont));
        }
        // ���������� �������
        for (String data : dataList) {
        	createFile(sheet, ++rowNum, table, myFont);          
        }
        // ������ ���������� ������ � pdf
		document.add(table);
		File excelFile = new File("Data.xls");
		String pathExcelFile = excelFile.getAbsolutePath();
		// ������ ���������� � ������ Excel ��������� � ����
        try (FileOutputStream out = new FileOutputStream(excelFile)) {
            workbook.write(out);   
            out.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("���� ������. ����: " + pathExcelFile);
        document.close();
        System.out.println("���� ������. ����: "+ pathPdfFile);
        writer.close();
    }
   
	// ��������� ����� � �������� ���������
    public static int randBetween(int start, int end) {
        return start + (int)Math.round(Math.random() * (end - start));
    }
    // ���������� �������� �� ���� ��������
    public static int calculateAge(LocalDate birthDate, LocalDate currentDate) {
        if ((birthDate != null) && (currentDate != null)) {
            return Period.between(birthDate, currentDate).getYears();
        } else {
            return 0;
        }
    }
    // ������ ���
    public static int[] calculateINN() {
	    int[] controlNumbers1 = {7, 2, 4, 10, 3, 5, 9, 4, 6, 8}; // ��������� ��� ����������� 1� ����������� �����
	    int[] controlNumbers2 = {3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8}; // ��������� ��� ����������� 2� ����������� �����
	    int[] inn = new int[12];
	    inn[0]=7; 
	    inn[1]=7;
	    inn[2]= randBetween(4, 0);
	    inn[3]= randBetween(9, 0);   
	    for(int i=4;i<10;i++)
	    {
	        inn[i]= randBetween(9, 0);
	    } 
	    int controlSum1 = 0;
	    for (int i = 0; i < 10; i++) {
	    	controlSum1 = controlSum1 + inn[i]*controlNumbers1[i];    			   	        
	    }
	    int rest1 = controlSum1 % 11;
	    if (rest1==10) {
	    	rest1 = 0;
	    }
	    inn[10] = rest1;
	    int controlSum2 = 0;
	    for (int i = 0; i < 11; i++) {
	    	controlSum2 = controlSum2 + inn[i]*controlNumbers2[i];    			   	        
	    }
	    int rest2 = controlSum2 % 11;
	    if (rest2==10) {
	    	rest2 = 0;
	    }
	    inn[11] = rest2;
	    return inn;
	}
    
    // ��������� ��������� ������ �� ���������� �����
    public static String getString(String fis) {
    List<String> fileLines = new ArrayList<String>();
     try
    {
   	 FileInputStream file =  new FileInputStream(fis);
        BufferedReader buffer = new BufferedReader(new InputStreamReader(file, "Cp1251"));
        String line = buffer.readLine();
            while (line != null)
            {
                fileLines.add(line);
                line = buffer.readLine();
            }
    }
    catch (Exception e)
    {
        e.printStackTrace();  
       
    } 
    Random randomStr = new Random();
    String data = fileLines.get(1+randomStr.nextInt(fileLines.size()-1));
    return data;
    }
    // ��������� ������ � ����������� �� ����
    public static List<String> getListData(String fis, String sex) {
        List<String> fileLinesMale = new ArrayList<String>();
        List<String> fileLinesFemale = new ArrayList<String>();
         try
        {
       	 FileInputStream file =  new FileInputStream(fis);
            BufferedReader buffer = new BufferedReader(new InputStreamReader(file, "Cp1251"));
            String line = buffer.readLine();
            while (line != null ) 
            {
            	if (line.charAt(line.length()-1)=='�'|| line.charAt(line.length()-1)=='�')
                {             	
            		fileLinesFemale.add(line);
                }
            	else {
            		fileLinesMale.add(line);
            	}
                    line = buffer.readLine();                         
            }  
        }
        catch (Exception e)
        {
            e.printStackTrace();  
        }     
         if (sex=="�") {
             return fileLinesFemale ; 
             } else {
          	   return fileLinesMale;
             }
    }
    
    public static String getData(List<String>fileLines) {
    	 Random randomStr = new Random();
         String data = fileLines.get(1+randomStr.nextInt(fileLines.size()-1));
         return data;
    } 
   
    // ���������� ������ (rowNum) ������������� ����� (sheet)
    // �������  �� c��������� � ������ Excel �����
    // � ������ ������ � pdf ���� 
        private static void createFile(HSSFSheet sheet, int rowNum, PdfPTable table, Font myFont) {
        Row row = sheet.createRow(rowNum);
        String name = getString("src\\main\\resources\\Name.txt"); //���
        String sex, surname, patronymic; //���, �������, ��������
        if (name.charAt(name.length()-1)=='�'|| name.charAt(name.length()-1)=='�') {
        	sex = "�";
        	 surname=getData(getListData("src\\main\\resources\\Surname.txt",sex));
        	 patronymic=getData(getListData("src\\main\\resources\\Patronymic.txt",sex));
        }
        else {
        	sex = "�";
        	surname=getData(getListData("src\\main\\resources\\Surname.txt",sex));
            patronymic=getData(getListData("src\\main\\resources\\Patronymic.txt",sex));
        }
        row.createCell(0).setCellValue(name);
        table.addCell(new Paragraph(name,myFont));
        row.createCell(1).setCellValue(surname);
        table.addCell(new Paragraph(surname,myFont));
        row.createCell(2).setCellValue(patronymic);
        table.addCell(new Paragraph(patronymic,myFont));
        Random randomBirthday = new Random();
        int minDay = (int) LocalDate.of(1920, 1, 1).toEpochDay();
        int maxDay = (int) LocalDate.of(2001, 1, 1).toEpochDay();
        long randomDay = minDay + randomBirthday.nextInt(maxDay - minDay);
        LocalDate randomBirthDate = LocalDate.ofEpochDay(randomDay); //��
        LocalDate randomBirthDateFormat = LocalDate.parse(randomBirthDate.toString(), DateTimeFormatter.ofPattern("yyyy-MM-dd")); //�� � ������ �������
        int age = calculateAge(randomBirthDate, LocalDate.now()); 
        row.createCell(3).setCellValue(age);
        table.addCell(new Paragraph(String.valueOf(age),myFont));
        row.createCell(4).setCellValue(sex);
        table.addCell(new Paragraph(sex,myFont));
        String BirthDate = randomBirthDateFormat.format(DateTimeFormatter.ofPattern("dd-MM-yyyy"));
        row.createCell(5).setCellValue(BirthDate);   
        table.addCell(new Paragraph(BirthDate,myFont));
        String inn = Arrays.toString(calculateINN()).replaceAll("\\[|\\]|,|\\s", "");
        row.createCell(6).setCellValue(inn);
        table.addCell(new Paragraph(inn,myFont));
        int randomPostCode = randBetween(200000, 100000);
        String postCode = String.valueOf(randomPostCode); //�������� ������
        row.createCell(7).setCellValue(postCode);     
        table.addCell(new Paragraph(String.valueOf(postCode),myFont));
        String country = getString("src\\main\\resources\\Country.txt"); //������
        row.createCell(8).setCellValue(country);
        table.addCell(new Paragraph(country,myFont));
        String area = getString("src\\main\\resources\\Area.txt");   //�������
        row.createCell(9).setCellValue(area);
        table.addCell(new Paragraph(area,myFont));
        String city = getString("src\\main\\resources\\City.txt");  //�����
        row.createCell(10).setCellValue(city);
        table.addCell(new Paragraph(city,myFont));
        String street = getString("src\\main\\resources\\Street.txt");  //�����
        row.createCell(11).setCellValue(street);
        table.addCell(new Paragraph(street,myFont));
        int randomHouse = randBetween(199, 1);
        String house = String.valueOf(randomHouse);  //���
        row.createCell(12).setCellValue(house);
        table.addCell(new Paragraph(house,myFont));
        int randomFlat = randBetween(999, 1); 
        String flat = String.valueOf(randomFlat); //��������
        row.createCell(13).setCellValue(flat);   
        table.addCell(new Paragraph(flat,myFont));
    }
        // ��������� ������ �������
        private static List<String> fillData() {
        List<String> data = new ArrayList<>();
        Random rNum = new Random();
        int rowNum = 1 + rNum.nextInt(29);
        System.out.println(rowNum);
        for (int i = 0;i < rowNum;i++) {
        	data.add(new String());
        }
        return data;
    }
}