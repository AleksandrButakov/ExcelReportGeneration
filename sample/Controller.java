package sample;

        import javafx.fxml.FXML;
        import javafx.scene.control.Button;
        import javafx.scene.control.TextArea;
        import javafx.scene.control.TextField;
        import javafx.scene.layout.AnchorPane;
        import javax.swing.*;
        import java.io.File;
        import java.io.IOException;


public class Controller {

    private String s1,s2,s3,s4;
    private int n1,n2,n3,n4;
    private int len, number_space;

    @FXML
    private AnchorPane formSample;

    @FXML
    private Button Button1;

    @FXML
    private Button Button2;

    @FXML
    private Button Button3;

    @FXML
    private TextField Text1;

    @FXML
    private TextField Text2;

    @FXML
    void initialize(){

        // нажатие на кнопку Veryfy
        Button1.setOnAction(event->{
            // проверка наличия файлов 18_1.xlsx, 18_2.xlsx, 18_3.xlsx, a.xlsx
            VerifyFile verifyFile = new VerifyFile();
            verifyFile.VerifyFile();
            if (VerifyFile.boolConditioCorrect == false) {
                return;
            }

            // открываем файл 18_1.xlsx, считываем контрольные поля, проверяем корректность
            // содержимого файлов, если все ОК, переносим данные в массив для дальнейшей обработки
            File1 file1 = new File1();
            try {
                file1.File1();
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (VerifyFile.boolConditioCorrect == false) {
                MessageFile1Error();
                return;
            }

            // открываем файл 18_2.xlsx, считываем контрольные поля, проверяем корректность
            // содержимого файлов, если все ОК, переносим данные в массив для дальнейшей обработки
            File2 file2 = new File2();
            try {
                file2.File2();
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (VerifyFile.boolConditioCorrect == false) {
                MessageFile1Error();
                return;
            }

            // открываем файл 18_3.xlsx, считываем контрольные поля, проверяем корректность
            // содержимого файлов, если все ОК, переносим данные в массив для дальнейшей обработки
            File3 file3 = new File3();
            try {
                file3.File3();
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (VerifyFile.boolConditioCorrect == false) {
                MessageFile1Error();
                return;
            }

            // открываем файл a.xlsx, считываем контрольные поля, проверяем корректность
            // содержимого файлов, если все ОК, переносим данные в массив для дальнейшей обработки
            FileA fileA = new FileA();
            try {
                fileA.FileA();
            } catch (IOException e) {
                e.printStackTrace();
            }
            if (VerifyFile.boolConditioCorrect == false) {
                MessageFile1Error();
                return;
            }

            // обработка полученных данных, формирование файла Result.xlsx
            JOptionPane.showMessageDialog(null, "Ok", "All very good!",
                    JOptionPane.INFORMATION_MESSAGE);

            ReportCreation reportCreation = new ReportCreation();
            reportCreation.ReportCreation();
            if (VerifyFile.boolConditioCorrect == false) {
                MessageFile1Error();
                return;
            }

            JOptionPane.showMessageDialog(null, "Класс report выполнен!", "All very good!",
                    JOptionPane.INFORMATION_MESSAGE);





        });


        // нажатие на кнопку 2
        Button2.setOnAction(event->{
            JOptionPane.showMessageDialog(null, "Null button", "Warning",
                    JOptionPane.ERROR_MESSAGE);
        });

        // нажатие на кнопку 3
        Button3.setOnAction(event->{

        });
    }

    public static void MessageFile1Error() {
        JOptionPane.showMessageDialog(null, "Выполнение программы завершено" +
                " по исключению!", "Warning", JOptionPane.ERROR_MESSAGE);
    }

}