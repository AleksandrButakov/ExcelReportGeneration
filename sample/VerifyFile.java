package sample;

import javax.swing.*;
import java.io.File;

public class VerifyFile {

    public static boolean boolConditioCorrect = true;

    public static void VerifyFile(){

        // проверка наличия файлов 18_1.xlsx, 18_2.xlsx, 18_3.xlsx, a.xlsx
        File file = new File("18_1.xlsx");
        if (file.exists() && file.isFile()){
            JOptionPane.showMessageDialog(null, "18_1.xlsx file is found", "Ok",
                    JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "18.1.xlsx file not found", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            boolConditioCorrect = false;
            return;
        }

        file = new File("18_2.xlsx");
        if (file.exists() && file.isFile()){
            JOptionPane.showMessageDialog(null, "18_2.xlsx file is found", "Ok",
                    JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "18.2.xlsx file not found", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            boolConditioCorrect = false;
            return;
        }

        file = new File("18_3.xlsx");
        if (file.exists() && file.isFile()){
            JOptionPane.showMessageDialog(null, "18_3.xlsx file is found", "Ok",
                    JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "18.3.xlsx file not found", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            boolConditioCorrect = false;
            return;
        }

        file = new File("a.xlsx");
        if (file.exists() && file.isFile()){
            JOptionPane.showMessageDialog(null, "a.xlsx file is found", "Ok",
                    JOptionPane.INFORMATION_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "a.xlsx file not found", "Warning",
                    JOptionPane.ERROR_MESSAGE);
            boolConditioCorrect = false;
            return;
        }

    }


}
