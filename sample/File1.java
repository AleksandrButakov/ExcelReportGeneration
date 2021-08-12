package sample;

import java.io.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.swing.*;

public class File1 {
    public static String arrStrFile1[][];
    public static int strokFile1=5;

    public static void File1() throws IOException {

        String strCell = "";
        // откроем файл 18_1.xlsx для чтения
        // откроем файл для чтения
        FileInputStream excelFile = new FileInputStream(new File("18_1.xlsx"));
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
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(1-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("№ п|п")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(2-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ФИО эксплуатационного штата, участвовавшего в работе")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(3-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("РВБ")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(4-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Всего трудозатрат")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(5-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ЛР Всего")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(6-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Инцидент")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(7-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ЛР И")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(8-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Запрос на изменение")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(9-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ЛР ЗИ")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(10-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Руководящее обращение")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(11-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ЛР РО")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(12-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("Обращение клиента")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

            row = sheet.getRow(5-1);
            cell = row.getCell(13-1);
            if (cell!=null) {
                strCell = cell.toString();
            }
            if (!strCell.equals("ЛР ОК")) {
                VerifyFile.boolConditioCorrect = false;
                MessageFile1Error();
                workbook.close();
                return;
            }

        } catch (java.lang.NullPointerException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(null, "Код класса File1 содержит ошибки!", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            VerifyFile.boolConditioCorrect = false;
            workbook.close();
            return;
        }

        // файл проверен, заголовок таблицы корректен. Определим количество строк в файле
        do {
            strokFile1++;
            try {
                row = sheet.getRow(strokFile1-1);
                cell = row.getCell(1-1);
            } catch (NullPointerException e) {
                // e.printStackTrace();
                break;
            }
        } while (cell!=null);
        strokFile1--;

        // зададим размерность массива String для файла 18_1.xlsx
        arrStrFile1 = new String[strokFile1-5][14];

        // последний столбец массива используем для хранения сокращенного ФИО
        // заполним массив данными
        for (int i=5; i<=strokFile1-1; i++) {
            for (int j=0; j<=12; j++) {
                row = sheet.getRow(i);
                cell = row.getCell(j);

                if (cell!=null) {
                    strCell = cell.toString();
                    arrStrFile1[i-5][j]=strCell;
                }
            }
        }

        // эта public переменная будет задавать размерность массива в других классах
        strokFile1=strokFile1-5;


        // заполним последний 13 столбец сокращенным ФИО, например Иванов И.И.
        String s1, s2;
        String surname, firstname, lastname;
        int len, n1, n2;
        for (int i=0; i<strokFile1; i++) {
            s1=arrStrFile1[i][1];
            len=s1.length();

            // проверим что ФИО содержит более 8 символов
            if (len<8) {
                // ФИО содержит менее 8 символов и введено не верно
                JOptionPane.showMessageDialog(null, "ФИО содержит менее 8 символов",
                        "Warning", JOptionPane.WARNING_MESSAGE);
                VerifyFile.boolConditioCorrect = false;
                workbook.close();
                return;
            }

            // проверим что ФИО содержит два пробела
            char ch;
            int count=0;
            for (int x=0; x<len; x++) {
                ch=s1.charAt(x);
                if (ch==' ') {
                    count++;
                }
            }
            if (count!=2) {
                JOptionPane.showMessageDialog(null, "Space != 2", "Error",
                        JOptionPane.WARNING_MESSAGE);
                VerifyFile.boolConditioCorrect = false;
                workbook.close();
                return;
            }

            // ФИО содержится в переменной s1 определим сокращенный формат
            n1=s1.indexOf(" ");
            n2=s1.indexOf(" ", n1+1);
            surname=s1.substring(0, n1);
            firstname=s1.substring(n1+1, n1+2);
            lastname=s1.substring(n2+1, n2+2);
            s2=surname + " " + firstname + ". " + lastname + ".";
            arrStrFile1[i][13]=s2;
        }


        workbook.close();
        JOptionPane.showMessageDialog(null, "Массив f1 заполнен!!!", "Ok",
                JOptionPane.INFORMATION_MESSAGE);
    }

    public static void MessageFile1Error() {
        JOptionPane.showMessageDialog(null, "Файл 18_1.xlsx содержит ошибки." +
                "Отчет не может быть сформирован.", "Warning", JOptionPane.ERROR_MESSAGE);
    }

}
