package sample;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Count;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.io.*;

public class ReportCreation {

    /* В этом классе создадим массив arrStrResult[][] в который поместим все данные из предыдудущих массивов
    * File1.arrStrFile1[strokFile1][14]
    * File2.arrStrFile2[strokFile2][6]
    * File3.arrStrFile3[strokFile1][6]
    * FileA.arrStrFileA[strokFileA][4]
    * В массиве arrStrResult[][] индексы столбцов будут содержать следующую информацию
    * 0 - Фамилия имя отчество
    * 1 - трудозатраты по ЛР И в часах
    * 2 - ЛР И количество
    * 3 - трудозатраты по ЛР ЗИ в часах
    * 4 - ЛР ЗИ количество
    * 5 - трудозатраты по ЛР РО в часах
    * 6 - ЛР РО количество
    * 7 - трудозатраты по ЛР ОК в часах
    * 8 - ЛР ОК количество
    * 9 - трудозатраты на ГТП-2 в часах (из File2 индекс столбца 3)
    * 10 - трудозатраты на совещания в часах (из File3 индекс столбца 3)
    * 11 - сумма трудозатрат по ЕСМА, будет формироваться из данных
    *      11 = 1 + 3 + 5 + 7 + 9 + 10 вычисляется программно
    * 12 - сокращенное значение ФИО, формируется программно из столбца 0
    * 13 - отработано часов всего (из FileA АСУТР, индекс столбца 12)   */


    public static void ReportCreation() {

        // зададим размерность массива (равна размерности АСУТР)
        String arrStrResult[][] = new String[FileA.strokFileA][14];
        int strokFileR = File1.strokFile1;
        String s1, s2;
        int i1, i2, i3, iA, iR;
        boolean bTemp;

        // заполним массив arrStrResult[][] значениями ""
        for (int y = 0; y < FileA.strokFileA; y++) {
            for (int x = 0; x < 14; x++) {
                arrStrResult[y][x] = "";
            }
        }

        // заполним arrStrResult[][] данными из arrStrFile1[][]
        for (i1 = 0; i1 < File1.strokFile1; i1++) {
            arrStrResult[i1][0] = File1.arrStrFile1[i1][1];
            arrStrResult[i1][1] = File1.arrStrFile1[i1][5];
            arrStrResult[i1][2] = File1.arrStrFile1[i1][6];
            arrStrResult[i1][3] = File1.arrStrFile1[i1][7];
            arrStrResult[i1][4] = File1.arrStrFile1[i1][8];
            arrStrResult[i1][5] = File1.arrStrFile1[i1][9];
            arrStrResult[i1][6] = File1.arrStrFile1[i1][10];
            arrStrResult[i1][7] = File1.arrStrFile1[i1][11];
            arrStrResult[i1][8] = File1.arrStrFile1[i1][12];
        }


        // добавим данные в arrStrResult[][9] из arrStrFile2[][3] 9 индекс массива
        for (i2 = 0; i2 < File2.strokFile2; i2++) {
            bTemp=false;
            for (iR = 0; iR < strokFileR; iR++) {
                s1 = File2.arrStrFile2[i2][1];
                s2 = arrStrResult[iR][0];
                if (s1.equals(s2)) {
                    // найдено соответствие фамилий, добавим трудозатраты ГТП-2
                    arrStrResult[iR][9] = File2.arrStrFile2[i2][3];
                    bTemp = true;
                }
            }
            if (bTemp == false) {
                arrStrResult[strokFileR][0] = File2.arrStrFile2[i2][1];
                arrStrResult[strokFileR][9] = File2.arrStrFile2[i2][3];
                strokFileR++;
            } else {
                bTemp = false;
            }
        }

        // добавим данные в arrStrResult[][10] из arrStrFile3[][] 10 индекс массива
        for (i3 = 0; i3 < File3.strokFile3; i3++) {
            bTemp = false;
            for (iR = 0; iR < strokFileR; iR++) {
                s1 = File3.arrStrFile3[i3][1];
                s2 = arrStrResult[iR][0];
                if (s1.equals(s2)) {
                    // найдено соответствие фамилий, добавим трудозатраты совещаний
                    arrStrResult[iR][10] = File3.arrStrFile3[i3][3];
                    bTemp = true;
                }
            }
            if (bTemp == false) {
                arrStrResult[strokFileR][0] = File3.arrStrFile3[i3][1];
                arrStrResult[strokFileR][10] = File3.arrStrFile3[i3][3];
                strokFileR++;
            } else {
                bTemp = false;
            }
        }
        strokFileR--;

        // рассчитаем значения столбца массива с индексом 11. 11 = 1 + 3 + 5 + 7 + 9 + 10
        float fSum, fT1, fT3, fT5, fT7, fT9, fT10;
        for (iR = 0; iR < strokFileR; iR++) {
            if (!arrStrResult[iR][1].equals("")) {   // && arrStrResult[iR][1] != null
                fT1 = Float.parseFloat(arrStrResult[iR][1]);
            } else {
                fT1 = 0f;
            }

            if (!arrStrResult[iR][3].equals("")) {
                fT3 = Float.parseFloat(arrStrResult[iR][3]);
            } else {
                fT3 = 0f;
            }

            if (!arrStrResult[iR][5].equals("")) {
                fT5 = Float.parseFloat(arrStrResult[iR][5]);
            } else {
                fT5 = 0f;
            }

            if (!arrStrResult[iR][7].equals("")) {
                fT7 = Float.parseFloat(arrStrResult[iR][7]);
            } else {
                fT7 = 0f;
            }

            if (!arrStrResult[iR][9].equals("")) {
                fT9 = Float.parseFloat(arrStrResult[iR][9]);
            } else {
                fT9 = 0f;
            }

            if (!arrStrResult[iR][10].equals("")) {
                fT10 = Float.parseFloat(arrStrResult[iR][10]);
            } else {
                fT10 = 0f;
            }

            fSum = fT1 + fT3 + fT5 + fT7 + fT9 + fT10;
            arrStrResult[iR][11] = String.valueOf(fSum);
        }


        // заполним столбец с индексом 12 сокращенным ФИО, например Иванов И.И.
        String surname, firstname, lastname;
        int n1, n2;
        for (iR = 0; iR < strokFileR; iR++) {
            s1 = arrStrResult[iR][0];

            // ФИО содержится в переменной s1 определим сокращенный формат
            n1 = s1.indexOf(" ");
            n2 = s1.indexOf(" ", n1+1);
            surname = s1.substring(0, n1);
            firstname = s1.substring(n1+1, n1+2);
            lastname=s1.substring(n2+1, n2+2);
            s2 = surname + " " + firstname + "." + lastname + ".";
            arrStrResult[iR][12] = s2;
        }

        char[] ch1;
        char[] ch2;

        // заполнис столбец с индексом 13 данными из АСУТР столбец: отработано часов, всего.
        // в массиве arrStrFileA столбец с индексом
        // добавим данные в arrStrResult[][9] из arrStrFile2[][3] 9 индекс массива
        for (iR = 0; iR < strokFileR; iR++) {
            bTemp=false;
            for (iA = 0; iA < FileA.strokFileA; iA++) {
                s1 = arrStrResult[iR][12];
                s2 = FileA.arrStrFileA[iA][0];
                ch1 = s1.toCharArray();
                ch2 = s2.toCharArray();

                if (s1.equals(s2)) {
                    // найдено соответствие фамилий, добавим трудозатраты ГТП-2
                    arrStrResult[iR][13] = FileA.arrStrFileA[iA][3];
                    bTemp = true;
                }
            }
            if (bTemp == false) {
                JOptionPane.showMessageDialog(null, "Не найдено соответствий ФИО ЕСМА и АСУТР." +
                        arrStrResult[iR][0], "Error", JOptionPane.WARNING_MESSAGE);
                VerifyFile.boolConditioCorrect = false;
                return;
            } else {
                bTemp = false;
            }
        }


        // создадим файл report.xlsx и перенесем в него данные из массива arrStrResult[][]
        Workbook workbookWrite = new XSSFWorkbook();
        Sheet sheetWrite = workbookWrite.createSheet("Report");
        Row rowWrite;
        Cell cellWrite;

        /*
        // запишем строку в новый файл Excel
        rowWrite = sheetWrite.createRow(2); //указываем номер столбца
        cellWrite = rowWrite.createCell(2, CellType.STRING);
        cellWrite.setCellValue("idquestions");
        cellWrite = rowWrite.createCell(5, CellType.STRING);
        cellWrite.setCellValue(5464);
        */

        /*
        for (int g = 2; g < 50; g++) {
                rowWrite = sheetWrite.createRow(g); //указываем номер столбца
            for (int h = 2; h < 50; h++) {
                cellWrite = rowWrite.createCell(h);
                cellWrite.setCellValue(h);
            }
        }
        */

        // сформируем заголовок
        //

        for (iR = 0; iR < strokFileR; iR++) {
                // переносим данные из массива в файл по ячейкам
                // номер столбца
                rowWrite = sheetWrite.createRow(iR+1);
                // номер строки
           for (int j = 0; j <= 13; j++) {
                cellWrite = rowWrite.createCell(j, CellType.BLANK);
                // значение ячейки

               try {
                   cellWrite.setCellValue(Float.parseFloat(arrStrResult[iR][j]));
               } catch (java.lang.NumberFormatException e) {
                   //e.printStackTrace();
                   cellWrite.setCellValue(arrStrResult[iR][j]);
               }

            }
        }

        // сохраним созданный файл
        FileOutputStream fos = null;
        try {
            fos = new FileOutputStream("temp_report.xlsx");
            workbookWrite.write(fos);
            fos.close();
            workbookWrite.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


/*
        try {
            FileOutputStream fosExcel = new FileOutputStream("Report.xlsx");
            Workbook workbookWrite = new XSSFWorkbook();
            Sheet sheetWrite = workbookWrite.createSheet("Report");
            Row rowWrite;
            Cell cellWrite;
            for (int y = 0; y < strokFileR; y++) {
                for (int x = 0; x <= 13; x++) {
                    // переносим данные из массива в файл по ячейкам
                    // номер столбца
                    rowWrite = sheetWrite.createRow(x);
                    // номер строки
                    cellWrite = rowWrite.createCell(y);
                    // значение ячейки
                    cellWrite.setCellValue(arrStrResult[y][x]);
                }
            }

            workbookWrite.write(fosExcel);
            fosExcel.close();
            workbookWrite.close();

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        // сохраним созданный файл
        //FileOutputStream fos = null;



         */
    }
}
