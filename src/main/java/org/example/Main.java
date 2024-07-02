package org.example;

import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.*;

import java.util.logging.Level;
import java.util.logging.Logger;

import static org.example.ConverterGUI.pathPreference;

public class Main {
    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());

    public static void main(String[] args) {

        try {
            // TODO: YOU CAN CHANG LOOK AND FEEL HERE
            // flatlaf: FlatLightLaf(), FlatDarkLaf()、FlatIntelliJLaf()
            UIManager.setLookAndFeel(new FlatLightLaf());
        } catch (UnsupportedLookAndFeelException e) {
            LOGGER.log(Level.SEVERE, "Unsupported Look and Feel", e);
        }

        String lastUsedFolderPath = pathPreference.getLastUsedFolder();
        String lastUsedFilePath = pathPreference.getLastUsedFilePath();

        SwingUtilities.invokeLater(() ->
                new ConverterGUI(lastUsedFolderPath, lastUsedFilePath).setVisible(true));
    }
}