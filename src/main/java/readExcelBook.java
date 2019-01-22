/**
 * @Author: EgbertW
 * @Description:
 * @Date: Created in 21:20 2019/1/9
 */
import java.awt.EventQueue;

import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.GroupLayout;
import javax.swing.JLabel;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JButton;
import javax.swing.JTextField;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import java.io.File;


public class readExcelBook {

    private JFrame frame;
    private JTextField textField;

    /**
     * Launch the application.
     */
    public static void main(String[] args) {
        EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    readExcelBook window = new readExcelBook();
                    window.frame.setVisible(true);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * Create the application.
     */
    public readExcelBook() {
        initialize();
    }

    /**
     * Initialize the contents of the frame.
     */
    private void initialize() {
        frame = new JFrame();
        frame.setBounds(100, 100, 450, 300);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        JButton button = new JButton("\u9009\u62E9\u6587\u4EF6");
        button.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc=new JFileChooser();
                jfc.setFileSelectionMode(JFileChooser.FILES_AND_DIRECTORIES );
                jfc.showDialog(new JLabel(), "选择");
                File file=jfc.getSelectedFile();
                String path = file.getAbsolutePath();
                textField.setText(path);
                ExcelRead er = new ExcelRead();
                er.readBook(path);
            }
        });

        textField = new JTextField();
        textField.setColumns(10);

        JLabel lbldaxls = new JLabel("\u5199\u5165\u4E86D\u76D8\u4E0B\u7684a.xls\u54C8");
        GroupLayout groupLayout = new GroupLayout(frame.getContentPane());
        groupLayout.setHorizontalGroup(
                groupLayout.createParallelGroup(Alignment.LEADING)
                        .addGroup(groupLayout.createSequentialGroup()
                                .addGap(26)
                                .addGroup(groupLayout.createParallelGroup(Alignment.LEADING)
                                        .addComponent(lbldaxls)
                                        .addComponent(textField, GroupLayout.PREFERRED_SIZE, 295, GroupLayout.PREFERRED_SIZE)
                                        .addComponent(button))
                                .addContainerGap(113, Short.MAX_VALUE))
        );
        groupLayout.setVerticalGroup(
                groupLayout.createParallelGroup(Alignment.LEADING)
                        .addGroup(groupLayout.createSequentialGroup()
                                .addGap(31)
                                .addComponent(button)
                                .addGap(18)
                                .addComponent(textField, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
                                .addGap(35)
                                .addComponent(lbldaxls)
                                .addContainerGap(119, Short.MAX_VALUE))
        );
        frame.getContentPane().setLayout(groupLayout);
    }
}
