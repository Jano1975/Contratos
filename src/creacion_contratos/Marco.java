/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package creacion_contratos;

import java.awt.*;
import javax.swing.*;

/**
 *
 * @author Usuario
 */
public class Marco extends JFrame {
    public Marco(){
        setTitle("Contratos");
        setBounds(500,100,600,680);
        JPanel lamina_cuadricula = new JPanel();
        lamina_cuadricula.setLayout(new GridLayout(10,3));
        
        
        
        Lamina lamina = new Lamina();
        Lamina lamina2 = new Lamina();
        
        lamina_cuadricula.add(lamina);
        lamina_cuadricula.add(lamina2);
        add(lamina_cuadricula); 
    }
    
}
