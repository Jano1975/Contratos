/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package creacion_contratos;

import java.awt.*;
import java.awt.event.ItemEvent;
import java.awt.event.ItemListener;
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
public class Marco extends JFrame {
    public Marco(){
        setTitle("Contratos");
        setBounds(500,300,800,680);
        JPanel lamina_cuadricula = new JPanel();
        lamina_cuadricula.setLayout(new GridLayout(10,3));
        EventoTipo tipo = new EventoTipo();
        lamina.comboBox.addItemListener(tipo);
        EventosDeRaton raton = new EventosDeRaton();
        lamina.boton.addMouseListener(raton);
        
        
        
        lamina_cuadricula.add(lamina);
        lamina_cuadricula.add(lamina2);
        lamina_cuadricula.add(lamina3);
        lamina2.setVisible(false);
        
        add(lamina_cuadricula); 
    }
    Lamina lamina = new Lamina();
    Lamina2 lamina2 = new Lamina2();
    Lamina3 lamina3 = new Lamina3();
    
     private class EventoTipo  implements ItemListener{

        @Override
        public void itemStateChanged(ItemEvent e) {
           if("Empresa".equals(e.getItem().toString())){
                lamina2.setVisible(true);
                lamina2.setVisible(true);
           }
           else{
                lamina2.setVisible(false);
                lamina2.setVisible(false);
                lamina2.tfRepresentante.setText("");
                lamina2.tfDNIRepresentante.setText("");
           }
        }
     
 }
 private class EventosDeRaton implements MouseListener{

    @Override
    public void mouseClicked(MouseEvent me) {
        
    }

    @Override
    public void mousePressed(MouseEvent me) {
      
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
            for(int c=0;c<10;c++){
                Cell celda = fila.createCell(c);
                if(f==0){
                    switch (c) {
                        case 0:
                            celda.setCellValue("DNI");
                            break;
                        case 1:
                            celda.setCellValue("Nombre");
                            break;
                        case 2:
                            celda.setCellValue("Apellidos");
                            break;
                        case 3:
                            celda.setCellValue("Dirección");
                            break;
                        case 4:
                            celda.setCellValue("Tipo");
                            break;
                        case 5:
                            celda.setCellValue("Representante");
                            break;
                        case 6:
                            celda.setCellValue("DNI Representante");
                            break;
                        case 7:
                            celda.setCellValue("Telefono");
                            break;
                        case 8:
                            celda.setCellValue("Telefono Móvil");
                            break;
                         case 9:
                            celda.setCellValue("Mail");
                            break;
                        default:
                            break;
                    }
                    
                }else{
                    switch (c) {
                        case 0:
                            celda.setCellValue(lamina.tfDNI.getText().trim());
                            break;
                        case 1:
                            celda.setCellValue(lamina.tfNombre.getText().trim());
                            break;
                        case 2:
                            celda.setCellValue(lamina.tfApellidos.getText().trim());
                            break;
                        case 3:
                            celda.setCellValue(lamina.tfDireccion.getText().trim());
                            break;
                        case 4:
                            celda.setCellValue(lamina.comboBox.getSelectedItem().toString());
                            break;
                        case 5: 
                            celda.setCellValue(lamina2.tfRepresentante.getText().trim());
                            break;
                        case 6:
                            celda.setCellValue(lamina2.tfDNIRepresentante.getText().trim());
                            break;    
                        case 7:
                            celda.setCellValue(lamina3.tfTelefono.getText().trim());
                            break;
                        case 8:
                            celda.setCellValue(lamina3.tfTelefonoMovil.getText().trim());
                            break;
                        case 9:
                            celda.setCellValue(lamina3.tfMail.getText().trim());
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

