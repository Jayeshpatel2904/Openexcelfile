package com.echostar;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.io.File;
import java.io.InputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings; // Corrected import
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;


/**
 * A Java Swing application to read specific sheets from a large Excel file
 * using a memory-efficient streaming (SAX) parser and display them in a JTable.
 *
 * @dependency Apache POI: You must have the Apache POI library JARs in your
 * project's classpath, including poi, poi-ooxml, and poi-ooxml-schemas.
 * You might also need xml-apis and xercesImpl. See pom.xml for Maven details.
 */
public class PopAnalysis extends JFrame {

    // The specific sheets we want to read from the Excel file.
    private static final String[] TARGET_SHEET_NAMES = {
        "Antenna_Electrical_Parameters",
        "Antennas",
        "NR_Sector_Carriers"
    };

    private JTabbedPane tabbedPane;
    private JLabel infoLabel;

    public PopAnalysis() {
        super("Excel Sheet Viewer (Streaming Mode)");

        // --- UI Setup ---
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1000, 700);
        setLocationRelativeTo(null); // Center the window

        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        JPanel topPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JButton openFileButton = new JButton("Open Excel File");
        infoLabel = new JLabel("Please select a large Excel file (.xlsx) to open.");
        topPanel.add(openFileButton);
        topPanel.add(infoLabel);

        tabbedPane = new JTabbedPane();
        mainPanel.add(topPanel, BorderLayout.NORTH);
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        add(mainPanel);

        // --- Action Listener for the Button ---
        openFileButton.addActionListener(e -> openAndProcessExcelFile());
    }

    /**
     * Opens a JFileChooser to select an Excel file and then processes it.
     */
    private void openAndProcessExcelFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Select an Excel File");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            public boolean accept(File f) {
                return f.getName().toLowerCase().endsWith(".xlsx") || f.isDirectory();
            }
            public String getDescription() {
                return "Excel Files (*.xlsx)";
            }
        });

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            infoLabel.setText("Processing: " + selectedFile.getName());
            tabbedPane.removeAll();
            // Run processing in a background thread to keep UI responsive
            new Thread(() -> processFile(selectedFile)).start();
        }
    }

    /**
     * Reads the selected Excel file using the memory-efficient SAX-based streaming parser.
     * @param file The Excel file to process.
     */
    private void processFile(File file) {
        try (OPCPackage pkg = OPCPackage.open(file)) {
            XSSFReader xssfReader = new XSSFReader(pkg);
            // Use the SharedStrings interface instead of the concrete class
            SharedStrings sst = xssfReader.getSharedStringsTable();
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter;

            boolean sheetsFound = false;
            // Create a new iterator for each sheet search because it can be consumed only once.
            for (String sheetName : TARGET_SHEET_NAMES) {
                iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData(); // Re-initialize iterator
                while (iter.hasNext()) {
                    try (InputStream stream = iter.next()) {
                        if (iter.getSheetName().equals(sheetName)) {
                            sheetsFound = true;
                            // Process this sheet
                            DefaultTableModel tableModel = new DefaultTableModel();
                            processSheet(new SheetToTableModelHandler(tableModel, sst), stream);
                            
                            // Update UI on the Event Dispatch Thread
                            SwingUtilities.invokeLater(() -> {
                                JTable table = new JTable(tableModel);
                                table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
                                table.setFillsViewportHeight(true);
                                JScrollPane scrollPane = new JScrollPane(table);
                                tabbedPane.addTab(sheetName, scrollPane);
                            });
                            break; // Move to the next target sheet name
                        }
                    }
                }
            }
            
            final boolean finalSheetsFound = sheetsFound;
            SwingUtilities.invokeLater(() -> {
                if (!finalSheetsFound) {
                    JOptionPane.showMessageDialog(this,
                        "None of the target sheets were found in the selected file.\n" +
                        "Required sheets: \n" + String.join("\n", TARGET_SHEET_NAMES),
                        "Sheets Not Found",
                        JOptionPane.WARNING_MESSAGE);
                    infoLabel.setText("No target sheets found. Please select another file.");
                } else {
                    infoLabel.setText("Successfully loaded: " + file.getName());
                }
            });

        } catch (Exception e) {
            SwingUtilities.invokeLater(() -> {
                JOptionPane.showMessageDialog(this,
                    "Error streaming the Excel file: " + e.getMessage(),
                    "File Read Error",
                    JOptionPane.ERROR_MESSAGE);
                infoLabel.setText("Error. Please try again.");
            });
            e.printStackTrace();
        }
    }

    /**
     * Parses a single sheet's XML data using a SAX parser.
     */
    private void processSheet(ContentHandler handler, InputStream sheetInputStream) throws IOException, SAXException {
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLReaderFactory.createXMLReader();
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (Exception e) {
            // This can happen with some Excel files, but we can often ignore it if data is read.
            System.err.println("Exception during parsing: " + e.getMessage());
            throw new RuntimeException("SAX parsing error", e);
        }
    }

    /**
     * Main method to run the application.
     */
    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }
            new PopAnalysis().setVisible(true);
        });
    }

    /**
     * The SAX handler to process sheet data row by row.
     */
    private static class SheetToTableModelHandler extends DefaultHandler {
        private final DefaultTableModel tableModel;
        // Use the SharedStrings interface
        private final SharedStrings sst;
        private String lastContents;
        private boolean nextIsString;
        private List<String> currentRow;
        private int headerColumnCount = 0;
        private boolean isHeaderRow = true;

        // Update constructor to accept the SharedStrings interface
        SheetToTableModelHandler(DefaultTableModel tableModel, SharedStrings sst) {
            this.tableModel = tableModel;
            this.sst = sst;
        }

        @Override
        public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
            // c => cell
            if (name.equals("c")) {
                // Figure out if the value is a string
                String cellType = attributes.getValue("t");
                nextIsString = (cellType != null && cellType.equals("s"));
            } else if (name.equals("row")) {
                currentRow = new ArrayList<>();
            }
            // Clear contents cache
            lastContents = "";
        }

        @Override
        public void endElement(String uri, String localName, String name) throws SAXException {
            // Process the content of a cell
            if (name.equals("c")) {
                String value = lastContents.trim();
                if (nextIsString) {
                    int idx = Integer.parseInt(value);
                    value = sst.getItemAt(idx).getString();
                }
                currentRow.add(value);
            } else if (name.equals("row")) {
                // We have reached the end of a row
                if (isHeaderRow && !currentRow.isEmpty()) {
                    for (String header : currentRow) {
                        tableModel.addColumn(header);
                    }
                    headerColumnCount = currentRow.size();
                    isHeaderRow = false;
                } else if (!isHeaderRow) {
                    // Pad row with empty strings if it's shorter than the header
                    while (currentRow.size() < headerColumnCount) {
                        currentRow.add("");
                    }
                    tableModel.addRow(currentRow.toArray());
                }
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) throws SAXException {
            lastContents += new String(ch, start, length);
        }
    }
}
