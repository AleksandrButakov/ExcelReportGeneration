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

public class FileA {
    public static String arrStrFileA[][];
    public static int strokFileA=22;
    public static String t;

    public static void FileA() throws IOException {

        String strCell = "";
        // откроем файл a.xlsx для чтения
        // откроем файл для чтения
        FileInputStream excelFile = new FileInputStream(new File("a.xlsx"));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet sheet = workbook.getSheetAt(0);
        Row row;
        Cell cell;

        // проведем проерку контрольных полей чтоб убедиться что файл корректен
        try {
            row = sheet.getRow(1-1);
            cell = row.getCell(4-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("                       Баланс рабочего времени")) {
                VerifyFile.boolConditioCorrect = false;
                t="88";
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(14-1);
            cell = row.getCell(1-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("№")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(23-1);
            cell = row.getCell(1-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("1.0")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(14-1);
            cell = row.getCell(2-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Фамилия,")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(14-1);
            cell = row.getCell(3-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Табельный")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(14-1);
            cell = row.getCell(4-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Штатная должность")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(14-1);
            cell = row.getCell(9-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Норма за")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

            row = sheet.getRow(16-1);
            cell = row.getCell(12-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Всего")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFileAError();
                workbook.close();
                return;
            }

        } catch (java.lang.NullPointerException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Код класса FileA содержит ошибки!", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            VerifyFile.boolConditioCorrect = false;
            workbook.close();
            return;
        }

        // файл проверен, заголовок таблицы корректен. Определим количество строк в файле
        do {
            strokFileA++;
            try {
                row = sheet.getRow(strokFileA-1);
                cell = row.getCell(1-1);
            } catch (NullPointerException e) {
                e.printStackTrace();
                break;
            }
        } while (!cell.equals("") || cell!=null);
        strokFileA--;
        strokFileA--;

        // зададим размерность массива String для файла A.xlsx
        arrStrFileA = new String[strokFileA-22][4];

        // заполним массив данными
        for (int i = 22; i < strokFileA; i++) {
            // заполним массив АСУТР данными

            row = sheet.getRow(i);
            cell = row.getCell(2-1);
            if (cell != null) {
                strCell = cell.toString();
                arrStrFileA[i-22][0]=strCell;
            }

            row = sheet.getRow(i);
            cell = row.getCell(4-1);
            if (cell!=null) {
                strCell = cell.toString();
                arrStrFileA[i-22][1]=strCell;
            }

            row = sheet.getRow(i);
            cell = row.getCell(9-1);
            if (cell!=null) {
                strCell = cell.toString();
                arrStrFileA[i-22][2]=strCell;
            }

            row = sheet.getRow(i);
            cell = row.getCell(12-1);
            if (cell!=null) {
                strCell = cell.toString();
                arrStrFileA[i-22][3]=strCell;
            }

        }

        // эта public переменная будет задавать размерность массива в других классах
        strokFileA=strokFileA-22;

        // заменим пробелы в нулевом столбце массива arrStrFileA[0][] в char кодировке
        // со 160 кода на 32. Необходимо для дальнейшего сравнения строк.
        char[] ch1;
        char[] ch2;
        String s1, s2;
        int len, ascii;
        for (int i = 0; i < strokFileA; i++) {
            s1 = arrStrFileA[i][0];
            ch1 = s1.toCharArray();
            len = s1.length();
            for (int j = 0; j < len; j++) {
                ascii = (int) ch1[j];
                // если код символа в массиве 160, меняем его на 32
                if (ascii == 160) {
                    ch1[j] = (char) 32;
                }
            }
            arrStrFileA[i][0] = String.valueOf(ch1);
        }


            workbook.close();
        JOptionPane.showMessageDialog(null, "Массив fa заполнен!!!", "Ok",
                JOptionPane.INFORMATION_MESSAGE);
    }

    public static void MessageFileAError() {
        JOptionPane.showMessageDialog(null, "Файл a.xlsx содержит ошибки." +
                "Отчет не может быть сформирован." + t, "Warning", JOptionPane.ERROR_MESSAGE);
    }


}
