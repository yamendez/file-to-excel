import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class EscribirExcel {
    private String[][] tabla;
    private HSSFWorkbook libro;
    private HSSFSheet hoja;
    private boolean temporal;

    public EscribirExcel(String[][] tabla, HSSFWorkbook libro, HSSFSheet hoja, boolean temporal) {
        this.tabla = tabla;
        this.libro = libro;
        this.hoja = hoja;
        this.temporal = temporal;
    }

    public void escribir(){
        String[] encabezado = tabla[0];
        int indiceFila = 0;
        HSSFRow encaFila = hoja.createRow(indiceFila);

        //Encabezado de tabla excel
        for(int i = 0; i < encabezado.length; i++){
            String header = encabezado[i];
            HSSFCell celda = encaFila.createCell(i);
            celda.setCellValue(header);
        }

        String[][] datos = new String[tabla.length-1][tabla[0].length];
        int val = 1;
        for(int i = 0; i < tabla.length-1; i++){
            datos[i] = tabla[val];
            val++;
        }

        indiceFila++;

        //Filas del Excel
        if(temporal){
            for(int i = 0; i < datos.length; i++){
                HSSFRow datarow = hoja.createRow(indiceFila);

                for(int j = 0; j < datos[i].length; j++){
                    if(datos[i][j].length()==16){
                        StringBuilder tempVal = new StringBuilder(datos[i][j]);
                        tempVal.setCharAt(1,'x');
                        datarow.createCell(j).setCellValue(tempVal.toString());
                    }else {
                        datarow.createCell(j).setCellValue(datos[i][j]);
                    }
                }
                indiceFila++;
            }
        }else {
            for (int i = 0; i < datos.length; i++) {
                HSSFRow datarow = hoja.createRow(indiceFila);

                for (int j = 0; j < datos[i].length; j++) {
                    datarow.createCell(j).setCellValue(datos[i][j]);
                }
                indiceFila++;
            }
        }

    }
}
