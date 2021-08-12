package sample;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class File3 {
    public static String arrStrFile3[][];
    public static int strokFile3=7;

    public static void File3() throws IOException {
        String strCell = "";
        // откроем файл 18_3.xlsx для чтения
        // откроем файл для чтения
        FileInputStream excelFile = new FileInputStream(new File("18_3.xlsx"));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
        Cell cell;

        // проведем проерку контрольных полей чтоб убедиться что файл корректен
        try {
            row = sheet.getRow(1-1);
            cell = row.getCell(1-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Название отчета - Распределение трудозатрат по сотрудникам")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(1-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("№ п|п")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(2-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ФИО эксплуатационного штата, участвовавшего в работе")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(3-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("РВБ")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(4-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("УК совещания, конференции(трудозатраты)")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(5-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Норма времени за период")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(7-1);
            cell = row.getCell(6-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Загрузка")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile3Error();
                workbook.close();
                return;
            }

        } catch (java.lang.NullPointerException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Код класса File3 содержит ошибки!", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            VerifyFile.boolConditioCorrect = false;
            workbook.close();
            return;
        }

        // файл проверен, заголовок таблицы корректен. Определим количество строк в файле
        do {
            strokFile3++;
            try {
                row = sheet.getRow(strokFile3-1);
                cell = row.getCell(1-1);
            } catch (NullPointerException e) {
                // e.printStackTrace();
                break;
            }
        } while (cell!=null);
        strokFile3--;
        strokFile3--;

        // зададим размерность массива String для файла 18_3.xlsx
        arrStrFile3 = new String[strokFile3-7][6];

        // заполним массив данными
        for (int i=7; i<=strokFile3-1; i++) {
            for (int j=0; j<=5; j++) {
                row = sheet.getRow(i);
                cell = row.getCell(j);

                if (cell!=null) {
                    strCell = cell.toString();
                    arrStrFile3[i-7][j]=strCell;
                }
            }
        }
        // эта public переменная будет задавать размерность массива в других классах
        strokFile3=strokFile3-7;

        workbook.close();
        JOptionPane.showMessageDialog(null, "Массив f3 заполнен!!!", "Ok",
                JOptionPane.INFORMATION_MESSAGE);
    }

    public static void MessageFile3Error() {
        JOptionPane.showMessageDialog(null, "Файл 18_3.xlsx содержит ошибки." +
                "Отчет не может быть сформирован.", "Warning", JOptionPane.ERROR_MESSAGE);
    }

}
