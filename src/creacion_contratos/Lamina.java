/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package creacion_contratos;

import java.awt.Desktop;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

/**
 *
 * @author Usuario
 */
public class Lamina extends JPanel {
    JLabel etNombre;
    JTextField tfNombre;
    JLabel etDireccion;
    JTextField tfDireccion;
    JComboBox comboBox;
    private String Tipo;
 public Lamina(){
     
     setBorder(BorderFactory.createTitledBorder(BorderFactory.createEtchedBorder(), "Datos"));
     
     setLayout(new BoxLayout(this, BoxLayout.X_AXIS));
     
     
         etNombre = new JLabel("Nombre");
         tfNombre = new JTextField();
         add(etNombre);
         add(tfNombre);
         comboBox = new JComboBox();
         comboBox.addItem("Escoge Tipo");
         comboBox.addItem("Persona");
         comboBox.addItem("Empresa");
         add(comboBox);
         etDireccion = new JLabel("Direccion");
         tfDireccion = new JTextField();
         add(etDireccion);
         add(tfDireccion);
             
     
     
     JButton boton = new JButton();
     boton.setText("Generar");
     add(boton);
     EventosDeRaton raton = new EventosDeRaton();
     boton.addMouseListener(raton);
     
     
     
     
     
 }   
 private class EventosDeRaton implements MouseListener{

    @Override
    public void mouseClicked(MouseEvent me) {
        
    }

    @Override
    public void mousePressed(MouseEvent me) {
      System.out.println("Has hecho click en el boton");
      String rutaArchivo = System.getProperty("user.home")+"/contratos.xls";
      File archivoXLS = new File(rutaArchivo);
      if(archivoXLS.exists()) archivoXLS.delete();

        try {
            archivoXLS.createNewFile();
        } catch (IOException ex) {
            Logger.getLogger(Lamina.class.getName()).log(Level.SEVERE, null, ex);
        }
        Workbook  libro = new HSSFWorkbook ();
        FileOutputStream archivo = null;
        try {
            archivo = new FileOutputStream(archivoXLS);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(Lamina.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        Sheet hoja =  libro.createSheet("Contratos");
        
        for(int f=0;f<2;f++){
            Row fila = hoja.createRow(f);
            for(int c=0;c<3;c++){
                Cell celda = fila.createCell(c);
                if(f==0){
                    switch (c) {
                        case 0:
                            celda.setCellValue("Nombre");
                            break;
                        case 1:
                            celda.setCellValue("Dirección");
                            break;
                        case 2:
                            celda.setCellValue("Tipo");
                            break;
                        default:
                            break;
                    }
                    
                }else{
                    switch (c) {
                        case 0:
                            celda.setCellValue(tfNombre.getText().trim());
                            break;
                        case 1:
                            celda.setCellValue(tfDireccion.getText().trim());
                            break;
                        case 2:
                            celda.setCellValue(comboBox.getSelectedItem().toString());
                            break;
                        default:
                            break;
                    }
                    
                }
            }
        }
        try {
            libro.write(archivo);
        } catch (IOException ex) {
            Logger.getLogger(Lamina.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        try {
            archivo.close();
        } catch (IOException ex) {
            Logger.getLogger(Lamina.class.getName()).log(Level.SEVERE, null, ex);
        }
        
        try {
            Desktop.getDesktop().open(archivoXLS);
        } catch (IOException ex) {
            Logger.getLogger(Lamina.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }

    @Override
    public void mouseReleased(MouseEvent me) {
        
    }

    @Override
    public void mouseEntered(MouseEvent me) {
        
    }

    @Override
    public void mouseExited(MouseEvent me) {
        
    }
    
}
}
