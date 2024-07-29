import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.xmlbeans.Filer;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

public class LeerArchivo {
    private File fileRead;
    private HSSFWorkbook libro;
    private HSSFSheet hoja;
    private String[][] tabla;
    private String nomTabla, separador;
    private boolean temporal;

    public LeerArchivo(File fileRead, HSSFWorkbook libro, HSSFSheet hoja, String nomTabla, String separador, boolean temporal) {
        this.fileRead = fileRead;
        this.libro = libro;
        this.hoja = hoja;
        this.nomTabla = nomTabla;
        this.separador = separador;
        this.temporal = temporal;
    }

    public void leer(){
        try {
            FileReader fr = new FileReader(fileRead);
            BufferedReader br = new BufferedReader(fr);

            br.mark(100000000);
            String linea = br.readLine();

            String[] data;
            int num_filas = 0;
            int num_column = 0;
            data = linea.split(new StringBuilder().append("\\").append(separador).toString());

            //Numero de columnas
            for(String value: data){
                num_column++;
            }

            //Número de filas
            while(linea != null){
                linea = br.readLine();
                num_filas++;
            }

            tabla = new String[num_filas][num_column];
            br.reset();
            linea = br.readLine();
            data = linea.split(new StringBuilder().append("\\").append(separador).toString());
            int i = 0;

            //Guardar los datos en un Arreglo
            while (linea != null) {
                data = linea.split(new StringBuilder().append("\\").append(separador).toString());

                for (int j = 0; j < tabla[0].length; j++) {
                    tabla[i][j] = data[j];
                }

                linea = br.readLine();
                i++;
            }

            System.out.println("filas:"+num_filas);
            System.out.println("columnas:" + num_column);

            EscribirExcel ee = new EscribirExcel(tabla, libro, hoja, temporal);
            ee.escribir();
            br.close();

        }catch (IOException ex){
            ex.printStackTrace();
        }
    }
}
