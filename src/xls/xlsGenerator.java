package xls;
import japa.parser.JavaParser;
import japa.parser.ast.CompilationUnit;
import japa.parser.ast.body.ClassOrInterfaceDeclaration;
import japa.parser.ast.body.FieldDeclaration;
import japa.parser.ast.body.MethodDeclaration;
import japa.parser.ast.visitor.VoidVisitorAdapter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.cli.*;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class xlsGenerator {

	/**
	 * @param args
	 */
	
	
	private static List<String> classInfo = new ArrayList<String>();
	private static List<List<String>> methods = new ArrayList<List<String>>();
	private static List<List<String>> fields = new ArrayList<List<String>>();
	
	private static void parseJava(File input) throws Exception{
		FileInputStream in = new FileInputStream(input);
		CompilationUnit cu;
        try {
            // parse the file
            cu = JavaParser.parse(in);
        } finally {
            in.close();
        }

        // visit and print the methods names
        new ClassVisitor().visit(cu, null);
        new MethodVisitor().visit(cu, null);
        new FieldVisitor().visit(cu, null);
    }

    /**
     * Simple visitor implementation for visiting MethodDeclaration nodes. 
     */
	private static class ClassVisitor extends VoidVisitorAdapter {

        @Override
        public void visit(ClassOrInterfaceDeclaration n, Object arg) {
            // here you can access the attributes of the method.
            // this method will be called for all methods in this 
            // CompilationUnit, including inner class methods
        	if(n.getJavaDoc()!=null){
        		classInfo.add(n.getName());
	        	String comment = n.getJavaDoc().getContent();
	        	comment = comment.replaceAll("\\* @param (.*)","");
	        	comment = comment.replaceAll("\\* @return (.*)","");
	        	comment = comment.replaceAll("\\* @throws (.*)","");
	        	comment = comment.replaceAll("\\* ","");
	        	comment = comment.replaceAll("(?s)\\*(.*)","");
	        	classInfo.add(comment.trim()); 
        	}
        	}
        }
	
	private static class MethodVisitor extends VoidVisitorAdapter {

        @Override
        public void visit(MethodDeclaration n, Object arg) {
            // here you can access the attributes of the method.
            // this method will be called for all methods in this 
            // CompilationUnit, including inner class methods
        	ArrayList<String> method = new ArrayList<String>();
        	if(n.getJavaDoc()!=null){
	        	method.add(n.getName());
	        	String comment = n.getJavaDoc().getContent();
	        	comment = comment.replaceAll("\\* @param (.*)","");
	        	comment = comment.replaceAll("\\* @return (.*)","");
	        	comment = comment.replaceAll("\\* @throws (.*)","");
	        	comment = comment.replaceAll("\\* ","");
	        	method.add(comment.trim()); 
	        	methods.add(method);
        	}
        	}
        }
    private static class FieldVisitor extends VoidVisitorAdapter {

        @Override
        public void visit(FieldDeclaration n, Object arg) {
            // here you can access the attributes of the method.
            // this method will be called for all methods in this 
            // CompilationUnit, including inner class methods
        	ArrayList<String> field = new ArrayList<String>();
        	if(n.getJavaDoc()!=null){
        		String declare = n.toString();
        		declare = declare.replaceAll("(?s)(.*)private", "private");
        		declare = declare.replaceAll("(?s)=(.*)", "");
        		field.add(declare.trim());
	        	String comment = n.getJavaDoc().getContent();
	        	comment = comment.replaceAll("\\* @param (.*)","");
	        	comment = comment.replaceAll("\\* @return (.*)","");
	        	comment = comment.replaceAll("\\* @throws (.*)","");
	        	comment = comment.replaceAll("\\* ","");
	        	field.add(comment.trim()); 
	        	fields.add(field);
        	}
        	}
        }
		
    private static void exportToExcel() throws IOException{
    	String filename = classInfo.get(0);
    	try {
    		// Create a New XL Document
    		HSSFWorkbook wb = new HSSFWorkbook();
    		// Make a worksheet in the XL document created
    		
    		HSSFSheet classsheet = wb.createSheet("class");
    		
    		int rowNum = 0; 

    		// Create the column headings 
    		HSSFRow headerRow = classsheet.createRow((short) rowNum); 
    		headerRow.createCell((short) 0).setCellValue(new HSSFRichTextString("Class Name:")); 
    		headerRow.createCell((short) 1).setCellValue(new HSSFRichTextString("Class Description:")); 
    		for(int i=0;i<1;i++){

    		// Create a new row in the sheet: 
    		HSSFRow row = classsheet.createRow((short) ++rowNum); 

    		HSSFCell nameCell = row.createCell((short) 0); 
    		nameCell.setCellValue(new HSSFRichTextString(classInfo.get(0))); 

    		HSSFCell infoCell = row.createCell((short) 1); 
    		infoCell.setCellValue(new HSSFRichTextString(classInfo.get(1))); 
    		}
    		HSSFSheet methodsheet = wb.createSheet("methods");
    		
    		rowNum = 0; 

    		// Create the column headings 
    		
    		HSSFRow headerRow1 = methodsheet.createRow((short) rowNum); 
    		headerRow1.createCell((short) 0).setCellValue(new HSSFRichTextString("Method Name:")); 
    		headerRow1.createCell((short) 1).setCellValue(new HSSFRichTextString("Method Description:")); 
    		for(int i=0;i<methods.size();i++){

    		// Create a new row in the sheet: 
    		HSSFRow row = methodsheet.createRow((short) ++rowNum); 

    		HSSFCell nameCell = row.createCell((short) 0); 
    		nameCell.setCellValue(new HSSFRichTextString(methods.get(i).get(0))); 

    		HSSFCell infoCell = row.createCell((short) 1); 
    		infoCell.setCellValue(new HSSFRichTextString(methods.get(i).get(1))); 
    		}
    		HSSFSheet fieldsheet = wb.createSheet("fields");
    		
    		rowNum = 0; 

    		// Create the column headings 
    		HSSFRow headerRow11 = fieldsheet.createRow((short) rowNum); 
    		headerRow11.createCell((short) 0).setCellValue(new HSSFRichTextString("Field Name:")); 
    		headerRow11.createCell((short) 1).setCellValue(new HSSFRichTextString("Field Description:")); 
    		for(int i=0;i<fields.size();i++){

    		// Create a new row in the sheet: 
    		HSSFRow row = fieldsheet.createRow((short) ++rowNum); 

    		HSSFCell nameCell = row.createCell((short) 0); 
    		nameCell.setCellValue(new HSSFRichTextString(fields.get(i).get(0))); 

    		HSSFCell infoCell = row.createCell((short) 1); 
    		infoCell.setCellValue(new HSSFRichTextString(fields.get(i).get(1))); 
    		}
    		
    		// The Output file is where the xls will be created
    		FileOutputStream fOut = new FileOutputStream(filename+".xls");
    		// Write the XL sheet
    		wb.write(fOut);
    		fOut.flush();
    		// Done Deal..
    		fOut.close();
    		System.out.println("File Created ..");
    		} catch (Exception e) {
    		System.out.println("exportToExcel() : " + e);
    		}

  		for(int i=0;i<methods.size();i++)
			for(int j=0;j<methods.get(i).size();j++)
			System.out.println(methods.get(i).get(j));
		for(int i=0;i<fields.size();i++)
			for(int j=0;j<fields.get(i).size();j++)
			System.out.println(fields.get(i).get(j));
		
    }
    private static File getFileFromArgs(String[] args) throws ParseException, IOException{
    	Options options = new Options();
    	
    	options.addOption("f", true, "local file path");
    	
    	CommandLineParser parser = new PosixParser();
    	CommandLine cmd = parser.parse( options, args);
    	
    	// get option values
    	String file = cmd.getOptionValue("f");

    	if(file != null) {
    		String filePath = file;
    		File javaFile = new File(filePath);
    		return javaFile;
    	}
    	else{
    		throw new IOException("Enter filename as arguement");
    	}
    	
    	
    }
	
	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File java = new File("BalanceInfo.java");
		parseJava(java);
		exportToExcel();
		
		
	}

}
