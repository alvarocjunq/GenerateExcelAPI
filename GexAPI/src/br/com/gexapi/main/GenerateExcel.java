package br.com.gexapi.main;

import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenerateExcel {
	
	/**
	 * Gerar Excel no caminho atual do projeto
	 * @param workbook Objeto do Excel
	 * @param nomeXLSSaida Nome do arquivo a ser gerado
	 */
	public static void gerarExcel(XSSFWorkbook workbook, String nomeXLSSaida) {
		try {
			FileOutputStream out = new FileOutputStream(nomeXLSSaida);
			workbook.write(out);
			out.flush();
			out.close();
//			workbook.close(); TODO: Avaliar a necessidade do fechamento
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * Adicionar uma aba (sheet)
	 * @param workbook XSSFWorkbook
	 * @param name Nome da sheet a ser criada no Excel
	 * @param lista Lista de registros a ser colocada na sheet (Objetos mapeados com a @PositionExcel 
	 * @param titles Titulos que vão aparecer no Header
	 * 
	 * @see {@link br.com.gexapi.main.PositionExcel}
	 */
	public static void addSheet(Workbook workbook, String name, List<?> lista, String...titles){
		Sheet sheet = workbook.createSheet(name);
		int qtdCol = setHeader(sheet, titles);
		generateLines(sheet, lista, qtdCol);
	}
	
	/**
	 * Gera as linhas no excel
	 * @param workbook Excel
	 * @param sheet Sheet aonde será escrita as linhas
	 * @param lista Lista de objetos anotados
	 * @param qtdCol Deve ser maior que a ultima posição colocada no objeto na @PositionExcel
	 * @param lineBegin Linha que deve começar a gerar as rows
	 */
	public static void generateLines(Sheet sheet, List<?> lista, int qtdCol, int lineBegin){
		for(int i=0,len=lista.size(); i<len;i++){
			setLine(sheet, lineBegin++, getCamposExcel(lista.get(i), qtdCol));
		}
	}
	
	private static void generateLines(Sheet sheet, List<?> lista, int qtdCol){
		for(int i=0,rownum=1,len=lista.size(); i<len;i++){
			setLine(sheet, rownum++, getCamposExcel(lista.get(i), qtdCol));
		}
	}
	
	private static void setLine(Sheet sheet, int rownum, String[] values){
		Row row;
		row = sheet.createRow(rownum);
		for(int i=0, len=values.length; i<len;i++){
			row.createCell(i).setCellValue(values[i]);
		}
	}
	
	/**
	 * Gera o Header da sheet
	 * @param workbook
	 * @param sheet
	 * @param titles Titulos a serem colocados no Header
	 * @return Retorna a quantidade de colunas geradas
	 */
	private static int setHeader(Sheet sheet, String[] titles){
		Row headerRow = sheet.createRow(0);
		int len=titles.length;
        for (int i=0; i < len; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(titles[i]);
        }
		return len;
	}
	
	
	private static String[] getCamposExcel(Object object, int qtdCol){
		String [] arrString = new String [qtdCol];
		try {
			//Itero sobre os metodos do objeto enviado
			for(Method f: object.getClass().getMethods()){
				//Verifica se tem a annotation PositionExcel
				if(f.isAnnotationPresent(PositionExcel.class)){
					
					//Pega as posicoes que a informacao deve ficar
					PositionExcel ann = f.getAnnotation(PositionExcel.class);
					int[] posicoes = ann.posicao();
					
					//Armazena no array de String, os valores dos campos
					//Se for necessário alterar o tipo do dado a ser colocado na célula, alterar nesse ponto
					for(int i=0,len=posicoes.length; i<len; i++){
						
						if(posicoes[i] > qtdCol)
							throw new RuntimeException("Número da coluna acessada, difere do números de colunas no HEADER\nQtd Header:"+qtdCol+" - Posição acessada:"+posicoes[i]);
						String value = String.valueOf(f.invoke(object, new Object[0]));
						arrString[posicoes[i]] = ("null".equals(value) ? "" : value);
					}
				}
			}
			return arrString;
		} catch ( SecurityException | IllegalAccessException | IllegalArgumentException | InvocationTargetException e1) {
			e1.printStackTrace();
			return new String[0];
		}
		
	}
	
}
