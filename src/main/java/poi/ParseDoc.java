package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ParseDoc {

	private Map<String, String> leis = new TreeMap<String, String>();		
	private static final String FILE_NAME_OUTPUT = "/home/francisco/Downloads/output/Leis.xlsx";	
	private static final String PATH_NAME_DOCS = "/home/francisco/Downloads/docs";

	public void parseDoc() throws IOException {
		
		XSSFWorkbook workbook = new XSSFWorkbook();

		File rootPath = new File(PATH_NAME_DOCS);
		File[] arquivos = rootPath.listFiles();

		for (File file : arquivos) {
			parse(file);
		}
		
		XSSFSheet sheet = null;
		int rowNum = 0;

		for (String s : leis.keySet()) {
			
			Pattern patternArt = Pattern.compile("Art.\\s?[0-9]{1,3}");
			Matcher matcherArt = patternArt.matcher(s);
			
			if(matcherArt.find()) {
				Row row = sheet.createRow(rowNum++);
				
				Cell cell = row.createCell(1);				
				cell.setCellValue(s);
				
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(leis.get(s));
				
			} else {
				sheet = workbook.createSheet(s);
				rowNum = 0;
				
				Row row = sheet.createRow(rowNum++);
				Cell cell = row.createCell(1);				
				cell.setCellValue("Artigo");
				
				Cell cell2 = row.createCell(2);
				cell2.setCellValue("Descrição");
				
			}
						
			System.out.println(s);
		}
					
		try {
            FileOutputStream outputStream = new FileOutputStream(FILE_NAME_OUTPUT);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
		
		System.out.println("Terminou");

	}

	public void parse(File file) throws IOException {

		FileInputStream is = new FileInputStream(file);

		WordExtractor extractor = new WordExtractor(is);

		String[] paragraphs = extractor.getParagraphText();

		boolean lei = true;

		String ultimaChave = "";
		String chave = "";
		String chaveLei = "";
				
		for (String paragraph : paragraphs) {

			Pattern patternLei = Pattern.compile("LEI N");
			Matcher matcherLei = patternLei.matcher(paragraph);

			Pattern patternArt = Pattern.compile("Art.\\s?[0-9]{1,3}");
			Matcher matcherArt = patternArt.matcher(paragraph);

			if (matcherLei.lookingAt() && lei) {
				int posicaoFinal = paragraph.indexOf(",");

				chave = paragraph.substring(0, posicaoFinal);

				ultimaChave = chave;
				chaveLei = chave;
				lei = false;

				leis.put(chave, paragraph);
			} else if (matcherArt.lookingAt()) {
				chave = paragraph.substring(0, 8);
				chave = chave.replace("º", "").trim();
				chave = chave.replace("°", "").trim();

				if (!chave.substring(4, 5).matches("\\s")) {
					String artigo = chave.substring(4, chave.length());

					chave = chave.substring(0, 4) + " " + artigo;
				}

				if (chave.endsWith(".")) {
					chave = chave.substring(0, chave.length() - 1);
				}

				if (chave.endsWith("-")) {
					chave = chave.substring(0, chave.length() - 1);
				}

				if (chave.substring(chave.length()-1).matches("[a-zA-Z]")) {
					chave = chave.substring(0, chave.length() - 1);
				}
				
				chave = chave.trim();
				
				
				if (chave.matches("Art.\\s[0-9]{2}")) {
					chave = chave.substring(0,5) + "0" + chave.substring(5,7);
				}else if (chave.matches("Art.\\s[0-9]{1}")) {
					chave = chave.substring(0,5) + "00" + chave.substring(5);
				}
								
				chave = chaveLei + " " + chave;
								
				if (leis.get(chave) != null) {

					if (chave.endsWith("1") && !chave.substring(chave.length()-2).matches("[1-9]")) {
						leis.put(chaveLei, leis.get(chaveLei) + leis.get(chave));
						leis.put(chave, paragraph);
					} else {
						leis.put(chaveLei, leis.get(chaveLei) + paragraph);
					}

				} else {
					leis.put(chave, paragraph);
				}

				ultimaChave = chave;

			} else {
				leis.put(ultimaChave, leis.get(ultimaChave) + paragraph);
			}

		}

		String[] lista = leis.get(ultimaChave).split("\\n");

		Pattern patternRodape = Pattern.compile("[0-9]{1,2}\\sde\\s[A-Za-zçÇ]{4,9}\\sde\\s[0-9]{4}");
		Matcher matcherRodape = null;
		int posicaoRodape = 0;

		for (int i = lista.length - 1; i >= 0; i--) {

			matcherRodape = patternRodape.matcher(lista[i]);

			if (matcherRodape.find()) {

				posicaoRodape = i;

				leis.put(chaveLei, leis.get(chaveLei) + "\n");

				for (int j = i; j < lista.length; j++) {
					leis.put(chaveLei, leis.get(chaveLei) + lista[j]);
				}

				break;

			}

		}

		if (posicaoRodape != 0) {

			leis.put(ultimaChave, "");

			for (int i = 0; i < lista.length; i++) {

				if (i == posicaoRodape) {
					break;
				}

				leis.put(ultimaChave, leis.get(ultimaChave) + lista[i]);
			}
		}
		
		extractor.close();

	}

}
