package program;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;

import entity.Aluno;

public class GenerateExcel {
	private static final String fileName = "C:/exemplos/alunos.xls";
	public static void main(String[] args) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheetAlunos = workbook.createSheet("Alunos");
		
		
		List<Aluno> listaAlunos = new ArrayList<>();
        listaAlunos.add(new Aluno("Eduardo", "9876525", 7, 8, 0, false));
        listaAlunos.add(new Aluno("Luiz", "1234466", 5, 8, 0, false));
        listaAlunos.add(new Aluno("Bruna", "6545657", 7, 6, 0, false));
        listaAlunos.add(new Aluno("Carlos", "3456558", 10, 3, 0, false));
        listaAlunos.add(new Aluno("Sonia", "6544546", 7, 8, 0, false));
        listaAlunos.add(new Aluno("Brianda", "3234535", 6, 5, 0, true));
        listaAlunos.add(new Aluno("Pedro", "4234524", 7, 5, 0, false));
        listaAlunos.add(new Aluno("Julio", "5434513", 7, 2, 0, false));
        listaAlunos.add(new Aluno("Henrique", "6543452", 7, 8, 0, true));
        listaAlunos.add(new Aluno("Fernando", "4345651", 5, 8, 0, false));
        listaAlunos.add(new Aluno("Vitor", "4332341", 7, 9, 0, false));
        

        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
        headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        CellStyle textStyle = workbook.createCellStyle();
        textStyle.setAlignment(CellStyle.ALIGN_CENTER);
        textStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        
        Cell cell;
        Row row;
        
        int rownum = 0;
        int cellnum = 0;
        
     // Configurando Header
        row = sheetAlunos.createRow(rownum++);
        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Nome");

        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("RA");

        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Nota 1");
        
        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Nota 2");
        
        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Média");
        
        cell = row.createCell(cellnum++);
        cell.setCellStyle(headerStyle);
        cell.setCellValue("Aprovado");

        
        for (Aluno aluno : listaAlunos) {
            row = sheetAlunos.createRow(rownum++);
            cellnum = 0;
            
            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(aluno.getNome());
            
            
            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(aluno.getRa());
            
            
            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(aluno.getNota1());
            
            
            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            cell.setCellValue(aluno.getNota2());
            
            
            cell = row.createCell(cellnum++);
            cell.setCellStyle(textStyle);
            Double media = aluno.getNota1() + aluno.getNota2() / 2;
            cell.setCellValue(media);
            
            
            Cell cellAprovado =row.createCell(cellnum++);
            cellAprovado.setCellValue(media >= 6);
            
            if(media >= 6) {
            	 row = sheetAlunos.createRow(rownum++);
            	 cellnum = 0;
            	 cell = row.createCell(cellnum++);
                 cell.setCellValue("by Filipe Souza Santos");
            }
        }
        
        try {
            FileOutputStream out = new FileOutputStream(new File(GenerateExcel.fileName));
            workbook.write(out);
            out.close();
            System.out.println("Arquivo Excel criado com sucesso!");

        } catch (FileNotFoundException e) {
            e.printStackTrace();
               System.out.println("Arquivo não encontrado!");
        } catch (IOException e) {
            e.printStackTrace();
               System.out.println("Erro na edição do arquivo!");
        }

	}
}
