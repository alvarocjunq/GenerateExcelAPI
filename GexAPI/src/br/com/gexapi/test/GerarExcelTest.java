package br.com.gexapi.test;

import static br.com.gexapi.main.GenerateExcel.addSheet;
import static br.com.gexapi.main.GenerateExcel.gerarExcel;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;

import br.com.gexapi.data.Cliente;

public class GerarExcelTest {

	private static final String PLANILHA_DE_CLIENTES_XLSX = "Planilha de clientes.xlsx";

	@Test
	public void generateExcel() {
		
		XSSFWorkbook workbook = new XSSFWorkbook();
		addSheet(workbook, "Clientes", getClientes(), "ID", "Nome");
		gerarExcel(workbook, PLANILHA_DE_CLIENTES_XLSX);
		
		Assert.assertTrue(new File(PLANILHA_DE_CLIENTES_XLSX).exists());
	}
	
	private static List<Cliente> getClientes(){
		List<Cliente> lista = new ArrayList<Cliente>();
		lista.add(new Cliente(1, "Joaquim"));
		lista.add(new Cliente(2, "João"));
		lista.add(new Cliente(3, "Joca"));
		lista.add(new Cliente(4, "Joaldo"));
		return lista;
	}

}
