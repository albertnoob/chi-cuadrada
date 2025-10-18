import java.io.File;
import java.io.FileInputStream;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ChiCuadrada  {

    public static void main(String[] args) {
        // Ruta del archivo Excel
        String archivo = System.getProperty("user.home") + "\\Downloads\\simulacionU2.xlsx";

        // Abrimos el archivo Excel usando Apache POI
        try (FileInputStream entrada = new FileInputStream(new File(archivo));
             XSSFWorkbook libro = new XSSFWorkbook(entrada)) {

            // Obtenemos la primera hoja del archivo
            Sheet hoja = libro.getSheetAt(0);

            // Formateador para leer valores como texto
            DataFormatter formato = new DataFormatter();

            // Tabla principal: edad → género → cantidad
            Map<String, Map<String, Integer>> tabla = new LinkedHashMap<>();

            //  guardar los géneros únicos
            Set<String> generos = new LinkedHashSet<>();

            // Recorremos las filas desde la segunda 
            for (int i = 1; i < hoja.getPhysicalNumberOfRows(); i++) {
                Row fila = hoja.getRow(i);
                if (fila == null) continue; // Ignora filas vacías

                // Extrae edad y género de cada fila
                String edad = formato.formatCellValue(fila.getCell(0));
                String genero = formato.formatCellValue(fila.getCell(1));

                // Agrega el género al conjunto si es nuevo
                generos.add(genero);

                // Si la edad no existe en la tabla, la crea
                tabla.putIfAbsent(edad, new LinkedHashMap<>());

                // Suma 1 al conteo de ese género en esa edad
                tabla.get(edad).put(genero, tabla.get(edad).getOrDefault(genero, 0) + 1);
            }

            // Muestra las tres tablas: observados, esperados y chi-cuadrado
            mostrarTodo(tabla, generos);

        } catch (Exception e) {
            // Manejo de errores al leer el archivo
            System.out.println("Ocurrió un error al leer el archivo:");
            e.printStackTrace();
        }
    }

    // Método que imprime las tres tablas con sus respectivas operaciones
    private static void mostrarTodo(Map<String, Map<String, Integer>> tabla, Set<String> generos) {
        // Calcula totales por edad, género y total general
        Map<String, Integer> totalEdad = new LinkedHashMap<>();
        Map<String, Integer> totalGenero = new LinkedHashMap<>();
        int totalGeneral = 0;

        for (String edad : tabla.keySet()) {
            int suma = 0;
            for (String genero : generos) {
                int val = tabla.get(edad).getOrDefault(genero, 0);
                suma += val;
                totalGenero.put(genero, totalGenero.getOrDefault(genero, 0) + val);
            }
            totalEdad.put(edad, suma);
            totalGeneral += suma;
        }
        System.out.println("\n==============================");
        System.out.println("TABLA DE DATOS");
        System.out.println("Operación: Conteo directo de Edad y Género");
        System.out.println("==============================");

        // Imprime encabezados
        System.out.printf("%-18s","Edad");
        for (String genero : generos) {
            System.out.printf("%-10s", genero);
        }
        System.out.printf("%-20s\n", "Total");

        // Imprime cada fila con conteo por edad y género
        for (String edad : tabla.keySet()) {
            System.out.printf("%-18s", edad);
            int filaTotal = 0;
            for (String genero : generos) {
                int val = tabla.get(edad).getOrDefault(genero, 0);
                filaTotal += val;
                System.out.printf("%-10d", val);
            }
            System.out.printf("%-18d\n", filaTotal);
        }

        // Imprime totales por género y total general
        System.out.printf("%-18s", "Total");
        for (String genero : generos) {
            System.out.printf("%-10d", totalGenero.getOrDefault(genero, 0));
        }
        System.out.printf("%-10d\n", totalGeneral);

        System.out.println("\n==============================");
        System.out.println("TABLA 2: VALORES ESPERADOS");
        System.out.println("Operación: (Total Edad * Total Género) / Total General");
        System.out.println("==============================");
     
        System.out.printf("%-18s","Edad");
        for (String genero : generos) {
            System.out.printf("%-10s", genero);
        }
        System.out.println();

        // Calcula e imprime valores esperados por celda
        for (String edad : tabla.keySet()) {
            System.out.printf("%-18s", edad);
            for (String genero : generos) {
                double esperado = (double) totalEdad.get(edad) * totalGenero.get(genero) / totalGeneral;
                System.out.printf("%-10.2f", esperado);
            }
            System.out.println();
        }

        System.out.println("\n==============================");
        System.out.println("TABLA CHI-CUADRADA");
        System.out.println("Operación: (Observado - Esperado)^2 / Esperado");
        System.out.println("==============================");
  
        // Imprime encabezados
        System.out.printf("%-18s","Edad");
        for (String genero : generos) {
            System.out.printf("%-10", genero);
        }
        System.out.println();

        // Calcula e imprime chi-cuadrado
        double chiTotal = 0;
        for (String edad : tabla.keySet()) {
            System.out.printf("%-18s", edad);
            for (String genero : generos) {
                int obs = tabla.get(edad).getOrDefault(genero, 0);
                double esp = (double) totalEdad.get(edad) * totalGenero.get(genero) / totalGeneral;
                double chi = Math.pow(obs - esp, 2) / esp;
                chiTotal += chi;
                System.out.printf("%-10.2f", chi);
            }
            System.out.println();
        }

        // Imprime el total general de chi-cuadrado
        System.out.printf("\nTotal Chi² = %.2f\n", chiTotal);

        
    }
    
}
