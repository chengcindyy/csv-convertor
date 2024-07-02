package org.example;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.util.List;
import java.util.concurrent.ExecutionException;

public class ConverterGUI extends JFrame {
    static final PathPreference pathPreference = new PathPreference();
    private static JTextField inputTextField;
    private static JTextField outputTextField;
    private static JLabel progressLabel;
    private static JProgressBar progressBar;

    private final TextProcessor textProcessor = new TextProcessor();

    ConverterGUI(String lastUsedFolderPath, String lastUsedFilePath) {
        // Set title and size
        setTitle("Csv to Xlsx Converter");
        setSize(370, 230);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setLayout(new BorderLayout());

        // Create components
        JPanel panel = new JPanel();
        JButton convertButton = new JButton("Convert to xlsx");
        JButton inputSelectButton = new JButton("Select Input File");
        JButton outputSelectButton = new JButton("Select Output Folder");
        JLabel headerLabel = new JLabel("CSV Convertor: ");
        JLabel contentLabel = new JLabel("Select a CSV File and a Output Folder");

        inputTextField = new JTextField(15);
        outputTextField = new JTextField(15);
        progressLabel = new JLabel("Click to start");
        progressBar = new JProgressBar();
        progressBar.setMinimum(0);
        progressBar.setMaximum(100);
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(322, 20));

        panel.add(headerLabel);
        panel.add(contentLabel);
        panel.add(inputSelectButton);
        panel.add(inputTextField);
        panel.add(outputSelectButton);
        panel.add(outputTextField);
        panel.add(progressBar);
        panel.add(convertButton);
        panel.add(progressLabel);

        add(panel, BorderLayout.CENTER);
        setVisible(true);
        // Use saved path and file name
        if (!lastUsedFolderPath.isEmpty()) {
            outputTextField.setText(lastUsedFolderPath);

        }
        if (!lastUsedFilePath.isEmpty()) {
            inputTextField.setText(lastUsedFilePath);
        }

        // Set button behavior
        inputSelectButton.addActionListener(e ->
                selectPathFromChooser(false, pathPreference.getLastUsedFilePath(), inputTextField, "csv"));

        outputSelectButton.addActionListener(e ->
                selectPathFromChooser(true, pathPreference.getLastUsedFolder(), outputTextField, null));

        // Call Convertor
        convertButton.addActionListener(e -> {
            String inputPath = inputTextField.getText().trim();
            String outputPath = outputTextField.getText().trim();

            convertButton.setEnabled(false);

            try {
                processDirectory(inputPath, outputPath);
            } finally {
                // Re-enable convert button after processing
                convertButton.setEnabled(true);
            }
        });

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

    private void selectPathFromChooser(boolean selectDirectory, String lastUsedPath, JTextField targetTextField, String fileTypeFilter) {
        JFileChooser chooser = createFileChooser(selectDirectory, lastUsedPath, fileTypeFilter);
        chooser.setDialogTitle(selectDirectory ? "Select Folder" : "Select " + fileTypeFilter + " File");

        if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = chooser.getSelectedFile();
            targetTextField.setText(selectedFile.getAbsolutePath());
            // Save the newly selected path
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

    private void processDirectory(String dirPath, String outputPath) {
        SwingWorker<Void, Integer> worker = new SwingWorker<>() {
            @Override
            protected Void doInBackground() throws Exception {
                File dir = new File(dirPath);
                if (dir.exists()) {
                    File[] fileList;
                    if (dir.isDirectory()) {
                        fileList = dir.listFiles();
                    } else {
                        fileList = new File[]{dir};
                    }

                    if (fileList != null && fileList.length > 0) {
                        int fileListSize = fileList.length;
                        System.out.println("Reading " + fileListSize + " file(s)...");

                        for (int i = 0; i < fileListSize; i++) {
                            if (isCancelled()) {
                                break; // Allow the swing worker to be cancellable
                            }

                            File file = fileList[i];
                            if (file.isFile()) {
                                String filePath = file.getAbsolutePath();
                                String extension = "";

                                int j = filePath.lastIndexOf('.');
                                if (j > 0) {
                                    extension = filePath.substring(j + 1).toLowerCase();
                                }

                                System.out.println("Processing " + file.getName() + "...");

                                if (extension.equals("csv")) {
                                    textProcessor.CSVToExcelProcessor(file, outputPath);
                                } else {
                                    System.err.println("Unsupported file type: " + extension);
                                }

                                // Update progress
                                int progress = (int) (((double) (i + 1) / fileListSize) * 100);
                                publish(progress);
                            }
                        }
                    } else {
                        JOptionPane.showMessageDialog(null, "No files found in the directory.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Invalid directory path.", "Error", JOptionPane.ERROR_MESSAGE);
                }
                return null;
            }

            @Override
            protected void process(List<Integer> chunks) {
                int mostRecentValue = chunks.get(chunks.size() - 1);
                progressBar.setValue(mostRecentValue);
                progressLabel.setText("Processing... " + mostRecentValue + "% completed");
            }

            @Override
            protected void done() {
                try {
                    get(); // Call get to rethrow exceptions occurred during doInBackground
                    System.out.println("Completed!");
                    progressBar.setValue(100);
                    progressLabel.setText("Completed!");
                    JOptionPane.showMessageDialog(null, "CSV file(s) processed successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);
                } catch (InterruptedException | ExecutionException e) {
                    e.printStackTrace();
                    progressLabel.setText("Failed!");
                    JOptionPane.showMessageDialog(null, "Error occurred during processing.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        };
        worker.execute();
    }
}