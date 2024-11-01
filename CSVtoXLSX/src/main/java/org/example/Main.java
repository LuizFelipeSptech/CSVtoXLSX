package org.example;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {

        Scanner scanner = new Scanner(System.in);
        System.out.print("Delimitador (Espaços seram considerados):");
        String delimiter = scanner.nextLine();

        File diretorio = new File("./");
        List<File> files = List.of(diretorio.listFiles());
        if(files == null || files.isEmpty()) {
            System.out.println("arquivos não encontrados, verifique se o arquivo 'CSVtoXLSX' está na mesma pasta dos seus arquivos");
            return;
        }
        for (File file : files) {
            if(file.getName().endsWith(".csv") && !file.getName().equalsIgnoreCase("CSVtoXLSX.jar")){
                try(BufferedReader br = new BufferedReader(new FileReader(file))){
                    List<List<String>> dados = new ArrayList<>();
                    String line;
                    while ((line = br.readLine()) != null) {
                        List<String> lineArray = Arrays.asList(line.split(delimiter));
                        dados.add(lineArray);
                    }

                    File writeenFile = tratarPath(file);
                    Boolean fileAlreadyExists = writeenFile.exists();

                    try (FileOutputStream out = new FileOutputStream(writeenFile)){
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet();
                        for (int rowIndex = 0; rowIndex < dados.size(); rowIndex++) {
                            Row row = sheet.createRow(rowIndex);
                            for (int cellIndex = 0; cellIndex < dados.get(rowIndex).size(); cellIndex++) {
                                row.createCell(cellIndex).setCellValue(dados.get(rowIndex).get(cellIndex));
                            }
                        }
                        if(fileAlreadyExists){
                            wb.write(out);
                            System.out.println("Arquivo recriado");
                        } else {
                            wb.write(out);
                            System.out.println("Arquivo criado");
                        }
                    }catch (Exception e){
                        System.out.println("Falha ao processar arquivo" + file.getName());
                    }

                }catch (Exception e){
                    e.printStackTrace();
                }
            } else {
                System.out.printf("Arquivo %s não é um csv\n", file.getName());
            }
        }
    }
    public static File tratarPath(File file){
        try{
            Files.createDirectories(Paths.get("dadosTratados"));
        } catch (Exception e) {
            e.printStackTrace();
        }
        String path = file.getPath();
        String nomeArquivo = path.substring(3,path.length()-4) + ".xlsx";
        System.out.printf("Tratando arquivo: %s\n", nomeArquivo);
        String pathCorreto = "dadosTratados/" + nomeArquivo;
        File writeenFile = new File(pathCorreto);
        return writeenFile;
    }
}