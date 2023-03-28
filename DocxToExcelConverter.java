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
        // ���� � ����� docx
        String filePath = "����.docx";
        // �������� ������ ������ �����
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("������");
        // ��������� ������ �� ��������� docx
        XWPFDocument document = new XWPFDocument(OPCPackage.open(new FileInputStream(filePath)));
        List<String> movieList = new ArrayList<String>();
        for (XWPFParagraph paragraph : document.getParagraphs()) {
            String text = paragraph.getText().trim();
            if (!text.isEmpty()) {
                // �������� ���������� � �������
                text = text.replaceAll("\\(.*?\\)", "");
                // ���������� �������� ������ � ������ �� ����� '-'
                String[] split = text.split(" - ");
                if (split.length == 2) {
                    movieList.add(split[0].trim());
                    movieList.add(split[1].trim());
                }
            }
        }
        // ������� ������ � ������� excel
        Iterator<String> iterator = movieList.iterator();
        int rowNum = 0;
        while (iterator.hasNext()) {
            Row row = sheet.createRow(rowNum++);
            Cell cell1 = row.createCell(0);
            cell1.setCellValue(iterator.next());
            Cell cell2 = row.createCell(1);
            cell2.setCellValue(Double.parseDouble(iterator.next()));
        }
        // ������ � ����� excel ����
        FileOutputStream outputStream = new FileOutputStream("����.xlsx");
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
    }
}