import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class DocxToExcelConverter {

    public static void main(String[] args) throws IOException, InvalidFormatException {
        // Путь к файлу docx
        String filePath = "файл.docx";
        // Создание нового эксель файла
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Фильмы");
        // Получение текста из документа docx
        XWPFDocument document = new XWPFDocument(OPCPackage.open(new FileInputStream(filePath)));
        List<String> movieList = new ArrayList<String>();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText().trim();
            if (!text.isEmpty()) {
                // Удаление информации в скобках
                text = text.replaceAll("\\(.*?\\)", "");
                // Разделение названия фильма и оценки по знаку '-'
                String[] split = text.split(" - ");
                if (split.length == 2) {
                    movieList.add(split[0].trim());
                    movieList.add(split[1].trim());
                }
            }
        }
        // Вставка данных в таблицу excel
        Iterator<String> iterator = movieList.iterator();
        int rowNum = 0;
        while (iterator.hasNext()) {
            Row row = sheet.createRow(rowNum++);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(iterator.next());
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(Double.parseDouble(iterator.next()));
        }
        // Запись в новый excel файл
        FileOutputStream outputStream = new FileOutputStream("файл.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
}