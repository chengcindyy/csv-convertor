package org.example;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.List;
import java.util.concurrent.ExecutionException;
import java.util.logging.Level;
import java.util.logging.Logger;

public class ConverterGUI extends JFrame {
    private static final Logger LOGGER = Logger.getLogger(Main.class.getName());

    private static JTextField inputTextField;
    private static JTextField outputTextField;
    private static JLabel progressLabel;
    private static JProgressBar progressBar;

    static final PathPreference pathPreference = new PathPreference();
    private static final TextProcessor textProcessor = new TextProcessor();

    public ConverterGUI(String lastUsedFolderPath, String lastUsedFilePath) {
        setTitle("Shopify Order Management System");
        setSize(370, 230);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.addTab("Step 1", createStep1Panel());
        tabbedPane.addTab("Step 2", createStep2Panel());
        add(tabbedPane, BorderLayout.CENTER);

        setVisible(true);

        if (!lastUsedFolderPath.isEmpty()) {
            outputTextField.setText(lastUsedFolderPath);
        }
        if (!lastUsedFilePath.isEmpty()) {
            inputTextField.setText(lastUsedFilePath);
        }
    }

    private JPanel createStep1Panel() {
        JPanel panel = new JPanel();
        inputTextField = new JTextField(15);
        outputTextField = new JTextField(15);
        progressLabel = new JLabel("Click to start");
        progressBar = createProgressBar();

        panel.add(new JLabel("CSV Convertor: "));
        panel.add(new JLabel("Select a CSV File and an Output Folder"));
        panel.add(createButton("Select Input File", _ -> selectPathFromChooser(false, pathPreference.getLastUsedFilePath(), inputTextField, "csv")));
        panel.add(inputTextField);
        panel.add(createButton("Select Output Folder", _ -> selectPathFromChooser(true, pathPreference.getLastUsedFolder(), outputTextField, null)));
        panel.add(outputTextField);
        panel.add(progressBar);
        panel.add(createButton("Convert to xlsx", _ -> processInputAndOutputFiles(inputTextField.getText().trim(), outputTextField.getText().trim())));
        panel.add(progressLabel);

        return panel;
    }

    private JPanel createStep2Panel() {
        JPanel panel = new JPanel();
        JTextField inputTextField = new JTextField(15);
        JLabel progressLabel = new JLabel("Click to start");
        JProgressBar progressBar = createProgressBar();

        panel.add(new JLabel("Import Tracking Number to Shopify"));
        panel.add(new JLabel("Select a File to Process"));
        panel.add(createButton("Select Input File", _ -> selectPathFromChooser(false, pathPreference.getLastUsedFilePath(), inputTextField, "xlsx")));
        panel.add(inputTextField);
        panel.add(progressBar);
        panel.add(createButton("Process", _ -> processInputFile(inputTextField.getText().trim())));
        panel.add(progressLabel);

        return panel;
    }

    private JProgressBar createProgressBar() {
        JProgressBar progressBar = new JProgressBar();
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(322, 20));
        return progressBar;
    }

    private JButton createButton(String text, java.awt.event.ActionListener actionListener) {
        JButton button = new JButton(text);
        button.addActionListener(actionListener);
        return button;
    }

    private void selectPathFromChooser(boolean selectDirectory, String lastUsedPath, JTextField targetTextField, String fileTypeFilter) {
        JFileChooser chooser = createFileChooser(selectDirectory, lastUsedPath, fileTypeFilter);
        chooser.setDialogTitle(selectDirectory ? "Select Folder" : "Select " + fileTypeFilter + " File");

        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = chooser.getSelectedFile();
            targetTextField.setText(selectedFile.getAbsolutePath());
            if (selectDirectory) {
                pathPreference.saveLastUsedFolder(selectedFile.getAbsolutePath());
            } else {
                pathPreference.saveLastUsedFilePath(selectedFile.getAbsolutePath());
            }
        } else {
            if (!lastUsedPath.isEmpty()) {
                targetTextField.setText(lastUsedPath);
            } else {
                System.out.println("No selection was made.");
            }
        }
    }

    private JFileChooser createFileChooser(boolean selectDirectory, String lastUsedPath, String fileTypeFilter) {
        JFileChooser chooser = new JFileChooser(lastUsedPath.isEmpty() ? new java.io.File(".") : new java.io.File(lastUsedPath));
        if (selectDirectory) {
            chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        } else {
            FileNameExtensionFilter filter = new FileNameExtensionFilter(fileTypeFilter + " Files", fileTypeFilter);
            chooser.setFileFilter(filter);
            chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        }
        chooser.setAcceptAllFileFilterUsed(false);
        return chooser;
    }

    private void processInputAndOutputFiles(String inputPath, String outputPath) {
        processFiles(inputPath, outputPath);
    }

    private void processInputFile(String inputPath) {
        processFiles(inputPath, "");
    }

    private void processFiles(String inputPath, String outputPath) {
        new FileProcessor(inputPath, outputPath).execute();
    }

    private class FileProcessor extends SwingWorker<Void, Integer> {
        private final String inputPath;
        private final String outputPath;

        public FileProcessor(String inputPath, String outputPath) {
            this.inputPath = inputPath;
            this.outputPath = outputPath;
        }

        @Override
        protected Void doInBackground() {
            File dir = new File(inputPath);
            if (dir.exists()) {
                File[] fileList = dir.isDirectory() ? dir.listFiles() : new File[]{dir};
                if (fileList != null && fileList.length > 0) {
                    int fileListSize = fileList.length;
                    for (int i = 0; i < fileListSize; i++) {
                        if (isCancelled()) break;
                        File file = fileList[i];
                        if (file.isFile()) {
                            String filePath = file.getAbsolutePath();
                            String extension = "";
                            int j = filePath.lastIndexOf('.');
                            if (j > 0) {
                                extension = filePath.substring(j + 1).toLowerCase();
                            }
                            if (extension.equals("csv")) {
                                textProcessor.CSVToExcelProcessor(file, outputPath);
                            } else if (extension.equals("xlsx")) {
                                ShopifyTracker.ImportDataProcessor(file);
                            } else {
                                System.err.println("Unsupported file type: " + extension);
                            }
                            int progress = (int) (((double) (i + 1) / fileListSize) * 100);
                            publish(progress);
                        }
                    }
                } else {
                    showErrorDialog("No files found in the directory.");
                }
            } else {
                showErrorDialog("Invalid directory path.");
            }
            return null;
        }

        @Override
        protected void process(List<Integer> chunks) {
            int mostRecentValue = chunks.getLast();
            progressBar.setValue(mostRecentValue);
            progressLabel.setText("Processing... " + mostRecentValue + "% completed");
        }

        @Override
        protected void done() {
            try {
                get();
                progressBar.setValue(100);
                progressLabel.setText("Completed!");
                showSuccessDialog();
            } catch (InterruptedException | ExecutionException e) {
                LOGGER.log(Level.SEVERE, "Process error", e);
                progressLabel.setText("Failed!");
                showErrorDialog("Error occurred during processing.");
            }
        }
    }

    private void showErrorDialog(String message) {
        JOptionPane.showMessageDialog(this, message, "Error", JOptionPane.ERROR_MESSAGE);
    }

    private void showSuccessDialog() {
        JOptionPane.showMessageDialog(this, "CSV file(s) processed successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> new ConverterGUI("", ""));
    }
}
