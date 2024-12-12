/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.teamlechuga.examenhugojose;

import java.awt.Component;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JSpinner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author GS2
 */
public class ExamenHugo extends javax.swing.JFrame {

    /**
     * Creates new form ExamenHugo
     */
    public ExamenHugo() {
        initComponents();
        setTitle("Notas Alumnos");
        setSize(600, 400);
        this.setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        iniciarPrograma();
    }
    
    private void iniciarPrograma(){
        Clase.addItem("1-A");
        Clase.addItem("1-B");
        String[] A = {"Hugo Jose", "Alejandro Moreno", "Marcos Plaza", "Agustin Camilo", "Roberto Perez", "Lucia Redondo", "Pepe Rodriguez", "Juan Marin"};
        String[] B = {"Mario Redondo", "Ruben Marin", "Alberto Manzanares", "Carlos Nazario", "Daniel Gonzalez", "Javier Lopez", "Carmelo Romero", "Camilo Andres"};
        Clase.addActionListener(e -> actualizarAlumnos(Clase, Alumnos, A, B));
        

        Calcular.addActionListener(e -> {
            calcularYExportar(Clase, Alumnos, MatematicasNota, LenguaNota, FisicaNota, TecnologiaNota);
        });
    }
    
    public static void actualizarAlumnos(JComboBox Clase, JComboBox Alumno, String[] A, String[] B){
        String claseSeleccionada = (String) Clase.getSelectedItem();

        Alumno.removeAllItems();

        if (claseSeleccionada.equals("1-A")) {
            for (String alumno : A) {
                Alumno.addItem(alumno);
            }
        } else if (claseSeleccionada.equals("1-B")) {
            for (String alumno : B) {
                Alumno.addItem(alumno);
            }
        }
    }
    
    public static void calcularYExportar(JComboBox ElegirClase, JComboBox ElegirAlumno, JSpinner MatematicasNota, JSpinner LenguaNota, JSpinner FisicaNota, JSpinner TecnologiaNota) {
        String nombreClase = (String) ElegirClase.getSelectedItem();
        String nombreAlumno = (String) ElegirAlumno.getSelectedItem();
        int notaMatematicas = (int) MatematicasNota.getValue();
        int notaLengua = (int) LenguaNota.getValue();
        int notaFisica = (int) FisicaNota.getValue();
        int notaTecnologia = (int) TecnologiaNota.getValue();


        if (nombreAlumno.isEmpty()) {
            JOptionPane.showMessageDialog(null, "El nombre del alumno es obligatorio.");
            return;
        }

        escribirEnExcel(nombreClase, nombreAlumno, notaMatematicas, notaLengua, notaFisica, notaTecnologia);
    }
    
    public static void escribirEnExcel(String clase, String alumno, int notaMatematicas, int notaLengua, int notaFisica, int notaTecnologia) {
        File archivoExcel = new File("C:\\Users\\GS2\\Desktop\\Notas Instituto.xlsx");
        Workbook libro = null;
        Sheet hoja = null;

        try {
            if (archivoExcel.exists()) {
                FileInputStream fis = new FileInputStream(archivoExcel);
                libro = new XSSFWorkbook(fis);
                hoja = libro.getSheet(clase);
                if (hoja == null) {
                    hoja = libro.createSheet(clase);
                }
                fis.close();
            } else {
                libro = new XSSFWorkbook();
                hoja = libro.createSheet(clase);
            }

            for (int i = hoja.getPhysicalNumberOfRows() - 1; i >= 0; i--) {
                Row row = hoja.getRow(i);
                if (row != null && row.getCell(0).getStringCellValue().equals("Medias")) {
                    hoja.removeRow(row);
                    break;
                }
            }

            if (hoja.getPhysicalNumberOfRows() == 0) {
                Row headerRow = hoja.createRow(0);
                headerRow.createCell(0).setCellValue("Nombre Alumno");
                headerRow.createCell(1).setCellValue("Nota Matematicas");
                headerRow.createCell(2).setCellValue("Nota Lengua");
                headerRow.createCell(3).setCellValue("Nota Fisica");
                headerRow.createCell(4).setCellValue("Nota Tecnologia");
            }

            boolean alumnoEncontrado = false;
            for (int i = 1; i < hoja.getPhysicalNumberOfRows(); i++) {
                Row row = hoja.getRow(i);
                if (row.getCell(0).getStringCellValue().equals(alumno)) {
                    row.getCell(1).setCellValue(notaMatematicas);
                    row.getCell(2).setCellValue(notaLengua);
                    row.getCell(3).setCellValue(notaFisica);
                    row.getCell(4).setCellValue(notaTecnologia);
                    alumnoEncontrado = true;
                    break;
                }
            }

            if (!alumnoEncontrado) {
                int nuevaFila = hoja.getPhysicalNumberOfRows();
                Row dataRow = hoja.createRow(nuevaFila);
                dataRow.createCell(0).setCellValue(alumno);
                dataRow.createCell(1).setCellValue(notaMatematicas);
                dataRow.createCell(2).setCellValue(notaLengua);
                dataRow.createCell(3).setCellValue(notaFisica);
                dataRow.createCell(4).setCellValue(notaTecnologia);
            }

            int ultimaFila = hoja.getPhysicalNumberOfRows();
            Row filaMedia = hoja.createRow(ultimaFila);

            filaMedia.createCell(0).setCellValue("Medias");

            double media = 0;
            int contar = 0;
            for (int col = 1; col <= 4; col++) {
                double suma = 0;
                int contador = 0;

                for (int i = 1; i < ultimaFila; i++) {
                    Row row = hoja.getRow(i);
                    if (row != null && row.getCell(col) != null) {
                        suma += row.getCell(col).getNumericCellValue();
                        contador++;
                        media += row.getCell(col).getNumericCellValue();
                        contar++;
                    }
                }

                double promedio = suma / contador;
                filaMedia.createCell(col).setCellValue(promedio);
            }

            double mediatotal = media / contar;
            filaMedia.createCell(5).setCellValue(mediatotal);

            for (int i = 0; i < 12; i++) {
                hoja.autoSizeColumn(i);
            }

            try (FileOutputStream fileOut = new FileOutputStream(archivoExcel)) {
                libro.write(fileOut);
                JOptionPane.showMessageDialog(null, "Datos exportados a Excel correctamente.");
            }
        } catch (IOException e) {
            JOptionPane.showMessageDialog(null, "Error al escribir en Excel: " + e.getMessage());
        } finally {
            try {
                if (libro != null) {
                    libro.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
    
    private void cambiarTamaño() {
        int tamanioSeleccionado = (Integer) MatematicasNota.getValue();

        for (Component comp : MatematicasNota.getComponents()) {
            if (comp instanceof textoPerso) {
                textoPerso lbl = (textoPerso) comp;
                lbl.cambiarTexto(tamanioSeleccionado, textoPerso1);
                System.out.println("Cambiando tamaño: " + tamanioSeleccionado);
            }
        }
    }


    

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {
        java.awt.GridBagConstraints gridBagConstraints;

        Clase = new javax.swing.JComboBox<>();
        Alumnos = new javax.swing.JComboBox<>();
        MatematicasTexto = new javax.swing.JLabel();
        MatematicasNota = new javax.swing.JSpinner();
        LenguaTexto = new javax.swing.JLabel();
        LenguaNota = new javax.swing.JSpinner();
        FisicaTexto = new javax.swing.JLabel();
        FisicaNota = new javax.swing.JSpinner();
        TecnologiaTexto = new javax.swing.JLabel();
        TecnologiaNota = new javax.swing.JSpinner();
        Calcular = new javax.swing.JButton();
        textoPerso1 = new com.teamlechuga.examenhugojose.textoPerso();
        MenuBarra = new javax.swing.JMenuBar();
        Menu = new javax.swing.JMenu();
        Notas1 = new javax.swing.JMenuItem();
        Tamaño = new javax.swing.JMenuItem();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(new java.awt.GridBagLayout());

        Clase.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " " }));
        Clase.setMinimumSize(new java.awt.Dimension(130, 26));
        Clase.setPreferredSize(new java.awt.Dimension(130, 26));
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(Clase, gridBagConstraints);

        Alumnos.setMinimumSize(new java.awt.Dimension(130, 26));
        Alumnos.setPreferredSize(new java.awt.Dimension(130, 26));
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(Alumnos, gridBagConstraints);

        MatematicasTexto.setText("Matematicas");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(MatematicasTexto, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 1;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.BOTH;
        getContentPane().add(MatematicasNota, gridBagConstraints);

        LenguaTexto.setText("Lengua");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(LenguaTexto, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 2;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.BOTH;
        getContentPane().add(LenguaNota, gridBagConstraints);

        FisicaTexto.setText("Fisica");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(FisicaTexto, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 3;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.BOTH;
        getContentPane().add(FisicaNota, gridBagConstraints);

        TecnologiaTexto.setText("Tecnologia");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 0;
        gridBagConstraints.gridy = 4;
        gridBagConstraints.gridwidth = 2;
        getContentPane().add(TecnologiaTexto, gridBagConstraints);
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 2;
        gridBagConstraints.gridy = 4;
        gridBagConstraints.gridwidth = 2;
        gridBagConstraints.fill = java.awt.GridBagConstraints.BOTH;
        getContentPane().add(TecnologiaNota, gridBagConstraints);

        Calcular.setText("Exportar");
        Calcular.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                CalcularActionPerformed(evt);
            }
        });
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 3;
        gridBagConstraints.gridy = 5;
        getContentPane().add(Calcular, gridBagConstraints);

        textoPerso1.setText("");
        gridBagConstraints = new java.awt.GridBagConstraints();
        gridBagConstraints.gridx = 4;
        gridBagConstraints.gridy = 1;
        getContentPane().add(textoPerso1, gridBagConstraints);

        Menu.setText("File");

        Notas1.setText("Poner notas a 1");
        Notas1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                Notas1ActionPerformed(evt);
            }
        });
        Menu.add(Notas1);

        Tamaño.setText("Tamaño de la ventana");
        Menu.add(Tamaño);

        MenuBarra.add(Menu);

        setJMenuBar(MenuBarra);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void CalcularActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_CalcularActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_CalcularActionPerformed

    private void Notas1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_Notas1ActionPerformed

    }//GEN-LAST:event_Notas1ActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ExamenHugo.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ExamenHugo.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ExamenHugo.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ExamenHugo.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ExamenHugo().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> Alumnos;
    private javax.swing.JButton Calcular;
    private javax.swing.JComboBox<String> Clase;
    private javax.swing.JSpinner FisicaNota;
    private javax.swing.JLabel FisicaTexto;
    private javax.swing.JSpinner LenguaNota;
    private javax.swing.JLabel LenguaTexto;
    private javax.swing.JSpinner MatematicasNota;
    private javax.swing.JLabel MatematicasTexto;
    private javax.swing.JMenu Menu;
    private javax.swing.JMenuBar MenuBarra;
    private javax.swing.JMenuItem Notas1;
    private javax.swing.JMenuItem Tamaño;
    private javax.swing.JSpinner TecnologiaNota;
    private javax.swing.JLabel TecnologiaTexto;
    private com.teamlechuga.examenhugojose.textoPerso textoPerso1;
    // End of variables declaration//GEN-END:variables
}
