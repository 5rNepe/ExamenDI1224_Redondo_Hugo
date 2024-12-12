/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.teamlechuga.examenhugojose;

import java.awt.Color;
import java.awt.Font;
import javax.swing.JLabel;
import javax.swing.SwingConstants;

/**
 *
 * @author GS2
 */
public class textoPerso extends JLabel {
    
    public textoPerso() {
        super("Texto");
        configurarEstilo();
    }

    public textoPerso(String texto) {
        super(texto);
        configurarEstilo();
    }

    private void configurarEstilo() {
        setHorizontalAlignment(SwingConstants.CENTER);
        setFont(new Font("Arial", Font.BOLD, 16));
    }

    public void cambiarTexto(int nota, JLabel texto) {
        if (nota >= 5) {
            setForeground(Color.BLACK);
            texto.setText("Aprobado");
        }
        else {
            setForeground(Color.RED);
            texto.setText("Suspenso");
        }
    }
}
