import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.net.URI;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class FileToExcel {
    private JPanel panelArchivo;
    private JTextField txtBuscar;
    private JButton btnBuscar;
    private JPanel panelTabla;
    private JCheckBox chbxTempFile;
    private JTextField txtNombre;
    private JTextField txtSeparador;
    private JButton btnEjecutar;
    private JPanel panelMain;

    private String direccion, nomTabla, separador;
    private File directorio;

    public FileToExcel() {
        btnBuscar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                JFileChooser fc = new JFileChooser();
                File dirActual = new File(System.getProperty("user.dir"));

                fc.setCurrentDirectory(dirActual);

                int seleccion = fc.showOpenDialog(fc);

                if(seleccion == JFileChooser.APPROVE_OPTION){
                    File fichero = fc.getSelectedFile();
                    txtBuscar.setText(fichero.getAbsolutePath());
                    direccion = fichero.toURI().toString();
                    directorio = fichero.getParentFile();

                }
            }
        });
        btnEjecutar.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                nomTabla = txtNombre.getText();
                separador = txtSeparador.getText();
                boolean temporal = chbxTempFile.isSelected();

                Date date = new Date();
                DateFormat df = new SimpleDateFormat("yyyyMMdd");

                FileWriter escribir = null;

                if(txtBuscar.getText().equals("") || nomTabla.equals("") || separador.equals("")){
                    JOptionPane.showMessageDialog(null,"Debe llenar todos los campos",
                            "Error Campos vacios", JOptionPane.ERROR_MESSAGE);
                }
                else {
                    HSSFWorkbook libro = new HSSFWorkbook();
                    HSSFSheet hoja = libro.createSheet();
                    libro.setSheetName(0,"Hoja");

                    File file;

                    if(temporal){
                        file = new File(directorio, nomTabla.toUpperCase() + "_TMP" + "_" + df.format(date)+ ".xls");
                    }else{
                        file = new File(directorio, nomTabla.toUpperCase() + "-" + df.format(date)+ ".xls");
                    }
                    try {
                        FileOutputStream fil = new FileOutputStream(file);

                        URI uri = new URI(direccion);
                        URL url = uri.toURL();

                        new LeerArchivo(new File(url.toURI()), libro, hoja, nomTabla, separador, temporal).leer();

                        libro.write(fil);
                        libro.close();

                        JOptionPane.showMessageDialog(null,"Operacion realizada con exito.",
                                "Mensaje", JOptionPane.INFORMATION_MESSAGE);

                        txtBuscar.setText("");
                        txtNombre.setText("");
                        txtSeparador.setText("");

                    } catch (Exception ex) {
                        throw new RuntimeException(ex);
                    }
                }
            }
        });
    }

    public static void main(String[] args){
        JFrame frame = new JFrame("File To Excel");
        frame.setContentPane(new FileToExcel().panelMain);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(400, 200);
        frame.setLocationRelativeTo(null);
        frame.pack();
        frame.setVisible(true);
    }
}
