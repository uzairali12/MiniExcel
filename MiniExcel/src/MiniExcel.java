import javax.swing.*;
import javax.swing.table.*;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.util.List;
import java.util.regex.*;

public class MiniExcel extends JFrame {
    private JTable table;
    private CustomTableModel model;
    private int rows = 45, cols = 13; // start with 13 cols, 45 rows
    private List<List<String>> sheet;
    private Stack<State> undoStack = new Stack<>();
    private Stack<State> redoStack = new Stack<>();
    private JTextField formulaBar;
    private JTable rowHeaderTable;
    private String clipboard = "";
    private boolean suppressStateChanges = false;
    private boolean showFormulas = false;

    private static class State {
        List<List<String>> sheet;
        int rows, cols;
        State(List<List<String>> sheet, int rows, int cols) {
            this.sheet = deepCopy(sheet);
            this.rows = rows;
            this.cols = cols;
        }
        private static List<List<String>> deepCopy(List<List<String>> original) {
            List<List<String>> copy = new ArrayList<>();
            for (List<String> row : original) copy.add(new ArrayList<>(row));
            return copy;
        }
    }

    public MiniExcel() {
        super("MiniExcel – A Spreadsheet Editor");
        try { UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName()); } catch(Exception ignored) {}

        initializeSheet();
        model = new CustomTableModel();

        String[] headers = new String[cols];
        for (int i = 0; i < cols; i++) headers[i] = getExcelColumnName(i);
        model.setColumnIdentifiers(headers);

        table = new JTable(model) {
            public String getToolTipText(MouseEvent e) {
                Point p = e.getPoint();
                int row = rowAtPoint(p), col = columnAtPoint(p);
                if(row >= 0 && col >= 0 && row < sheet.size() && col < sheet.get(row).size()) {
                    String raw = sheet.get(row).get(col);
                    return raw != null && raw.startsWith("=") ? raw : raw;
                }
                return null;
            }
        };
        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        table.setSelectionMode(ListSelectionModel.SINGLE_SELECTION);
        table.setCellSelectionEnabled(true);
        table.setShowGrid(true);
        table.setGridColor(Color.GRAY);
        table.setIntercellSpacing(new Dimension(1,1));
        table.setRowHeight(24);
        table.getTableHeader().setReorderingAllowed(false);
        table.getTableHeader().setBackground(new Color(220,220,220));
        table.getTableHeader().setFont(new Font("Arial", Font.BOLD, 12));

        updateColumnWidths(); // will be adjusted after adding to scroll pane

        rowHeaderTable = new JTable(new RowHeaderModel());
        rowHeaderTable.setRowHeight(table.getRowHeight());
        rowHeaderTable.setPreferredScrollableViewportSize(new Dimension(50,0));
        rowHeaderTable.setSelectionModel(table.getSelectionModel());
        rowHeaderTable.setColumnSelectionAllowed(false);
        rowHeaderTable.setRowSelectionAllowed(false);
        rowHeaderTable.setBackground(new Color(230,230,230));
        rowHeaderTable.setForeground(Color.DARK_GRAY);
        rowHeaderTable.setFont(new Font("Arial", Font.BOLD, 12));
        rowHeaderTable.setBorder(BorderFactory.createLineBorder(Color.BLACK));

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setRowHeaderView(rowHeaderTable);
        scrollPane.setCorner(JScrollPane.UPPER_LEFT_CORNER, new Corner() {
            @Override public Dimension getPreferredSize() { return new Dimension(50, table.getRowHeight()); }
        });
        scrollPane.setHorizontalScrollBarPolicy(JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);

        updateColumnWidths(); // after scroll pane

        JPanel formulaPanel = new JPanel(new BorderLayout());
        JLabel formulaLabel = new JLabel("Formula: ");
        formulaLabel.setFont(new Font("Arial", Font.BOLD, 12));
        formulaPanel.add(formulaLabel, BorderLayout.WEST);
        formulaBar = new JTextField();
        formulaBar.setBackground(Color.LIGHT_GRAY);
        formulaBar.setFont(new Font("Arial", Font.PLAIN, 12));
        formulaPanel.add(formulaBar, BorderLayout.CENTER);

        add(formulaPanel, BorderLayout.NORTH);
        add(scrollPane, BorderLayout.CENTER);
        setJMenuBar(createMenuBar());

        JToolBar bottomToolBar = createToolBar();
        bottomToolBar.setBackground(new Color(255,200,0));
        add(bottomToolBar, BorderLayout.SOUTH);

        DefaultCellEditor cellEditor = new DefaultCellEditor(new JTextField()) {
            @Override
            public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row, int column) {
                JTextField editor = (JTextField) super.getTableCellEditorComponent(table, value, isSelected, row, column);
                if(row >= 0 && column >=0 && row<sheet.size() && column<sheet.get(row).size()) {
                    String raw = sheet.get(row).get(column);
                    editor.setText(raw==null?"":raw);
                }
                return editor;
            }
            @Override
            public boolean stopCellEditing() {
                boolean ok = super.stopCellEditing();
                if(ok) {
                    int row = table.getEditingRow(), col = table.getEditingColumn();
                    if(row>=0 && col>=0) model.setRawValueAt((String)getCellEditorValue(), row, col);
                }
                return ok;
            }
        };
        table.setDefaultEditor(Object.class, cellEditor);

        table.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                if(e.isControlDown()) {
                    switch(e.getKeyCode()) {
                        case KeyEvent.VK_C: copyCell(); break;
                        case KeyEvent.VK_V: pasteCell(); break;
                        case KeyEvent.VK_X: cutCell(); break;
                        case KeyEvent.VK_Z: undo(); break;
                        case KeyEvent.VK_Y: redo(); break;
                        case KeyEvent.VK_F2:
                            int r = table.getSelectedRow(), c = table.getSelectedColumn();
                            if(r>=0 && c>=0) table.editCellAt(r,c);
                            break;
                    }
                }
            }
        });

        table.getSelectionModel().addListSelectionListener(e -> updateFormulaBar());
        table.getColumnModel().getSelectionModel().addListSelectionListener(e -> updateFormulaBar());

        formulaBar.addActionListener(e -> {
            int r = table.getSelectedRow(), c = table.getSelectedColumn();
            if(r>=0 && c>=0) model.setRawValueAt(formulaBar.getText(), r, c);
        });

        saveState();

        setDefaultCloseOperation(EXIT_ON_CLOSE);
        setSize(1400,800);
        setVisible(true);
    }

    private JMenuBar createMenuBar() {
        JMenuBar menuBar = new JMenuBar();

        // File Menu
        JMenu fileMenu = new JMenu("File");
        JMenuItem saveItem = new JMenuItem("Save");
        saveItem.addActionListener(e -> saveCSV());
        JMenuItem loadItem = new JMenuItem("Load");
        loadItem.addActionListener(e -> loadCSV());
        JMenuItem exitItem = new JMenuItem("Exit");
        exitItem.addActionListener(e -> System.exit(0));
        fileMenu.add(saveItem);
        fileMenu.add(loadItem);
        fileMenu.addSeparator();
        fileMenu.add(exitItem);
        menuBar.add(fileMenu);

        // Edit Menu
        JMenu editMenu = new JMenu("Edit");
        JMenuItem undoItem = new JMenuItem("Undo");
        undoItem.addActionListener(e -> undo());
        JMenuItem redoItem = new JMenuItem("Redo");
        redoItem.addActionListener(e -> redo());
        JMenuItem cutItem = new JMenuItem("Cut");
        cutItem.addActionListener(e -> cutCell());
        JMenuItem copyItem = new JMenuItem("Copy");
        copyItem.addActionListener(e -> copyCell());
        JMenuItem pasteItem = new JMenuItem("Paste");
        pasteItem.addActionListener(e -> pasteCell());
        editMenu.add(undoItem);
        editMenu.add(redoItem);
        editMenu.addSeparator();
        editMenu.add(cutItem);
        editMenu.add(copyItem);
        editMenu.add(pasteItem);
        menuBar.add(editMenu);

        // Insert Menu
        JMenu insertMenu = new JMenu("Insert");
        JMenuItem insertRowItem = new JMenuItem("Insert Row");
        insertRowItem.addActionListener(e -> insertRow());
        JMenuItem insertColItem = new JMenuItem("Insert Column");
        insertColItem.addActionListener(e -> insertColumn());
        insertMenu.add(insertRowItem);
        insertMenu.add(insertColItem);
        menuBar.add(insertMenu);

        // Delete Menu
        JMenu deleteMenu = new JMenu("Delete");
        JMenuItem deleteRowItem = new JMenuItem("Delete Row");
        deleteRowItem.addActionListener(e -> deleteRow());
        JMenuItem deleteColItem = new JMenuItem("Delete Column");
        deleteColItem.addActionListener(e -> deleteColumn());
        deleteMenu.add(deleteRowItem);
        deleteMenu.add(deleteColItem);
        menuBar.add(deleteMenu);

        // View menu with Show Formulas (kept for completeness)
        JMenu viewMenu = new JMenu("View");
        JCheckBoxMenuItem showFormMenuItem = new JCheckBoxMenuItem("Show Formulas");
        showFormMenuItem.addActionListener(e -> {
            showFormulas = showFormMenuItem.isSelected();
            model.fireTableDataChanged();
        });
        viewMenu.add(showFormMenuItem);
        menuBar.add(viewMenu);

        return menuBar;
    }

    private JToolBar createToolBar() {
        JToolBar toolBar = new JToolBar();
        JButton sumBtn = new JButton("SUM");
        sumBtn.setFont(new Font("Arial", Font.BOLD, 12));
        sumBtn.addActionListener(e -> insertFunction("SUM"));
        JButton avgBtn = new JButton("AVG");
        avgBtn.setFont(new Font("Arial", Font.BOLD, 12));
        avgBtn.addActionListener(e -> insertFunction("AVG"));
        JButton meanBtn = new JButton("MEAN");
        meanBtn.setFont(new Font("Arial", Font.BOLD, 12));
        meanBtn.addActionListener(e -> insertFunction("MEAN"));
        JButton minBtn = new JButton("MIN");
        minBtn.setFont(new Font("Arial", Font.BOLD, 12));
        minBtn.addActionListener(e -> insertFunction("MIN"));
        JButton maxBtn = new JButton("MAX");
        maxBtn.setFont(new Font("Arial", Font.BOLD, 12));
        maxBtn.addActionListener(e -> insertFunction("MAX"));
        JButton countBtn = new JButton("COUNT");
        countBtn.setFont(new Font("Arial", Font.BOLD, 12));
        countBtn.addActionListener(e -> insertFunction("COUNT"));
        JButton medianBtn = new JButton("MEDIAN");
        medianBtn.setFont(new Font("Arial", Font.BOLD, 12));
        medianBtn.addActionListener(e -> insertFunction("MEDIAN"));
        JButton modeBtn = new JButton("MODE");
        modeBtn.setFont(new Font("Arial", Font.BOLD, 12));
        modeBtn.addActionListener(e -> insertFunction("MODE"));
        JButton stdevBtn = new JButton("STDEV");
        stdevBtn.setFont(new Font("Arial", Font.BOLD, 12));
        stdevBtn.addActionListener(e -> insertFunction("STDEV"));
        JButton rangeBtn = new JButton("RANGE");
        rangeBtn.setFont(new Font("Arial", Font.BOLD, 12));
        rangeBtn.addActionListener(e -> insertFunction("RANGE"));
        JButton productBtn = new JButton("PRODUCT");
        productBtn.setFont(new Font("Arial", Font.BOLD, 12));
        productBtn.addActionListener(e -> insertFunction("PRODUCT"));
        JButton absBtn = new JButton("ABS");
        absBtn.setFont(new Font("Arial", Font.BOLD, 12));
        absBtn.addActionListener(e -> insertFunction("ABS"));
        JButton sqrtBtn = new JButton("SQRT");
        sqrtBtn.setFont(new Font("Arial", Font.BOLD, 12));
        sqrtBtn.addActionListener(e -> insertFunction("SQRT"));

        toolBar.add(sumBtn);
        toolBar.add(avgBtn);
        toolBar.add(meanBtn);
        toolBar.add(minBtn);
        toolBar.add(maxBtn);
        toolBar.add(countBtn);
        toolBar.addSeparator();
        toolBar.add(medianBtn);
        toolBar.add(modeBtn);
        toolBar.add(stdevBtn);
        toolBar.add(rangeBtn);
        toolBar.add(productBtn);
        toolBar.addSeparator();
        toolBar.add(absBtn);
        toolBar.add(sqrtBtn);
        return toolBar;
    }
    private void updateColumnWidths() {
        TableColumnModel cm = table.getColumnModel();
        for (int i = 0; i < cm.getColumnCount(); i++) {
            TableColumn col = cm.getColumn(i);
            col.setPreferredWidth(120); // adjust width as needed
            col.setMinWidth(60);
        }
    }
    private void initializeSheet() {
        sheet = new ArrayList<>();
        for (int i = 0; i < rows; i++) {
            List<String> row = new ArrayList<>();
            for (int j = 0; j < cols; j++) {
                row.add("");
            }
            sheet.add(row);
        }
    }

    // Converts 0-based column index to Excel-style name
// 0 -> A, 25 -> Z, 26 -> AA, 27 -> AB, 701 -> ZZ, 702 -> AAA
    private String getExcelColumnName(int index) {
        StringBuilder name = new StringBuilder();
        index++; // convert to 1-based

        while (index > 0) {
            int rem = (index - 1) % 26;
            name.insert(0, (char) ('A' + rem));
            index = (index - 1) / 26;
        }
        return name.toString();
    }

    private void saveState() {
        if (suppressStateChanges) return;
        // push a snapshot to undo stack
        undoStack.push(new State(sheet, rows, cols));
        // new change -> clear redo
        redoStack.clear();
    }

    private void undo() {
        if (!undoStack.isEmpty()) {
            suppressStateChanges = true;
            // push current to redo
            redoStack.push(new State(sheet, rows, cols));
            State prev;
            prev = undoStack.pop();
            sheet = prev.sheet;
            rows = prev.rows;
            cols = prev.cols;
            refreshTable();
            suppressStateChanges = false;
        } else {
            Toolkit.getDefaultToolkit().beep();
        }
    }

    private void redo() {
        if (!redoStack.isEmpty()) {
            suppressStateChanges = true;
            undoStack.push(new State(sheet, rows, cols));
            State next = redoStack.pop();
            sheet = next.sheet;
            rows = next.rows;
            cols = next.cols;
            refreshTable();
            suppressStateChanges = false;
        } else {
            Toolkit.getDefaultToolkit().beep();
        }
    }

    private void refreshTable() {
        int selRow = table.getSelectedRow();
        int selCol = table.getSelectedColumn();

        model.setRowCount(rows);
        model.setColumnCount(cols);

        // Update column headers
        String[] headers = new String[cols];
        for (int i = 0; i < cols; i++) {
            headers[i] = getExcelColumnName(i);
        }
        model.setColumnIdentifiers(headers);

        // ❌ DO NOT destroy structure
        model.fireTableDataChanged();
        updateColumnWidths();
        table.doLayout();
        // Re-sync row header
        rowHeaderTable.revalidate();
        rowHeaderTable.repaint();

        // Restore selection safely
        if (rows > 0 && cols > 0) {
            if (selRow < 0) selRow = 0;
            if (selCol < 0) selCol = 0;
            selRow = Math.min(selRow, rows - 1);
            selCol = Math.min(selCol, cols - 1);

            table.setRowSelectionInterval(selRow, selRow);
            table.setColumnSelectionInterval(selCol, selCol);
        }

        updateFormulaBar();
    }

    private void insertRow() {
        saveState();
        rows++;
        List<String> newRow = new ArrayList<>();
        for (int j = 0; j < cols; j++) {
            newRow.add("");
        }
        sheet.add(newRow);
        refreshTable();
    }

    private void deleteRow() {
        if (rows > 1) {
            saveState();
            rows--;
            sheet.remove(sheet.size() - 1);
            refreshTable();
        }
    }

    private void insertColumn() {
        saveState();
        cols++;
        for (List<String> row : sheet) {
            row.add("");
        }
        refreshTable();
    }

    private void deleteColumn() {
        if (cols > 1) {
            saveState();
            cols--;
            for (List<String> row : sheet) {
                row.remove(row.size() - 1);
            }
            refreshTable();
        }
    }

    private void saveCSV() {
        try (PrintWriter pw = new PrintWriter(new File("sheet.csv"))) {
            for (List<String> row : sheet) {
                StringBuilder sb = new StringBuilder();
                for (int j = 0; j < row.size(); j++) {
                    String cell = row.get(j) == null ? "" : row.get(j);
                    // escape any commas by wrapping in quotes if necessary
                    if (cell.contains(",") || cell.contains("\"") || cell.contains("\n")) {
                        cell = "\"" + cell.replace("\"", "\"\"") + "\"";
                    }
                    sb.append(cell);
                    if (j < row.size() - 1) sb.append(",");
                }
                pw.println(sb.toString());
            }
            JOptionPane.showMessageDialog(this, "Saved to sheet.csv");
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "Error: " + e.getMessage());
        }
    }

    private void loadCSV() {
        try(BufferedReader br = new BufferedReader(new FileReader("sheet.csv"))) {
            List<List<String>> loaded = new ArrayList<>();
            String line;
            while((line=br.readLine())!=null) {
                List<String> parsed = parseCSVLine(line);
                loaded.add(parsed);
            }
            if(!loaded.isEmpty()) {
                sheet = loaded;
                rows = sheet.size();
                cols = sheet.get(0).size();
                // pad missing cells
                for(List<String> row:sheet) while(row.size()<cols) row.add("");
            } else initializeSheet();
            saveState();
            refreshTable();
        } catch(FileNotFoundException fnf) {
            JOptionPane.showMessageDialog(this, "sheet.csv not found.");
        } catch(Exception e) {
            JOptionPane.showMessageDialog(this, "Error: "+e.getMessage());
        }
    }

    // Simple CSV parser for quotes
    private List<String> parseCSVLine(String line) {
        List<String> result = new ArrayList<>();
        if (line == null || line.isEmpty()) return result;
        StringBuilder cur = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < line.length() && line.charAt(i + 1) == '"') {
                        cur.append('"');
                        i++; // skip next
                    } else {
                        inQuotes = false;
                    }
                } else {
                    cur.append(c);
                }
            } else {
                if (c == '"') {
                    inQuotes = true;
                } else if (c == ',') {
                    result.add(cur.toString());
                    cur.setLength(0);
                } else {
                    cur.append(c);
                }
            }
        }
        result.add(cur.toString());
        return result;
    }

    private void updateFormulaBar() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row >= 0 && col >= 0 && row < sheet.size() && col < sheet.get(row).size()) {
            formulaBar.setText(sheet.get(row).get(col));
        } else {
            formulaBar.setText("");
        }
    }

    private void insertFunction(String func) {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row >= 0 && col >= 0) {
            String range = JOptionPane.showInputDialog("Enter args (e.g., A1:B2 or A1,10,B1:C2):");
            if (range != null) {
                String formula = "=" + func + "(" + range + ")";
                formulaBar.setText(formula);
                model.setRawValueAt(formula, row, col);
            }
        }
    }

    private void copyCell() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row >= 0 && col >= 0) {
            clipboard = sheet.get(row).get(col);
        }
    }

    private void pasteCell() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row >= 0 && col >= 0) {
            model.setRawValueAt(clipboard, row, col);
        }
    }

    private void cutCell() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row >= 0 && col >= 0) {
            clipboard = sheet.get(row).get(col);
            model.setRawValueAt("", row, col);
        }
    }

    // Row Header Model
    private class RowHeaderModel extends AbstractTableModel {
        @Override
        public int getRowCount() {
            return rows;
        }

        @Override
        public int getColumnCount() {
            return 1;
        }

        @Override
        public Object getValueAt(int rowIndex, int columnIndex) {
            return rowIndex + 1;
        }

        @Override
        public boolean isCellEditable(int rowIndex, int columnIndex) {
            return false;
        }
    }

    // Corner for row header
    private class Corner extends JComponent {
        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            g.setColor(Color.LIGHT_GRAY);
            g.fillRect(0, 0, getWidth(), getHeight());
            g.setColor(Color.BLACK);
            g.drawString("Row", 5, 15);
        }
    }

    // Custom Table Model
    private class CustomTableModel extends DefaultTableModel {
        @Override
        public Object getValueAt(int row, int column) {
            if (row < sheet.size() && column < sheet.get(row).size()) {
                String raw = sheet.get(row).get(column);
                if (showFormulas) {
                    // return raw content directly for Show Formulas mode
                    return raw == null ? "" : raw;
                }
                if (raw != null && raw.startsWith("=")) {
                    double result = evaluateFormula(raw);
                    if (Double.isNaN(result)) {
                        return "ERROR";
                    }
                    if (result == (long) result) {
                        return String.valueOf((long) result);
                    } else {
                        return String.format("%.2f", result);
                    }
                }
                return raw == null ? "" : raw;
            }
            return "";
        }

        @Override
        public void setValueAt(Object aValue, int row, int column) {
            if (row < sheet.size() && column < sheet.get(row).size()) {
                // Save state before modifying so user can undo this change
                saveState();
                sheet.get(row).set(column, aValue == null ? "" : aValue.toString());
                fireTableCellUpdated(row, column);
            }
        }

        public void setRawValueAt(String value, int row, int column) {
            if (row < sheet.size() && column < sheet.get(row).size()) {
                saveState();
                sheet.get(row).set(column, value == null ? "" : value);
                fireTableCellUpdated(row, column);
            }
        }
    }

    /**
     * Evaluate formula string (starting with '=')
     * Supports functions and nested references and arithmetic expressions
     */
    private double evaluateFormula(String expr) {
        try {
            expr = expr.substring(1).trim(); // remove '=' and trim spaces

            // process functions iteratively until no functions remain
            expr = processAllFunctions(expr);

            // Replace single cell references like A1, B2, AA10 etc. with their numeric values (or 0)
            Pattern pattern = Pattern.compile("([A-Z]+)([0-9]+)");
            Matcher matcher = pattern.matcher(expr);
            StringBuffer sb = new StringBuffer();

            while (matcher.find()) {
                String token = matcher.group();
                int[] cell = parseCell(token);
                String val = "0";
                if (cell != null && cell[0] >= 0 && cell[0] < rows && cell[1] >= 0 && cell[1] < cols) {
                    String raw = sheet.get(cell[0]).get(cell[1]);
                    if (raw != null && raw.startsWith("=")) {
                        double nested = evaluateFormula(raw);
                        if (!Double.isNaN(nested)) {
                            val = String.valueOf(nested);
                        } else {
                            val = "0";
                        }
                    } else {
                        if (raw == null || raw.isEmpty()) {
                            val = "0";
                        } else {
                            try {
                                Double.parseDouble(raw);
                                val = raw;
                            } catch (NumberFormatException e) {
                                val = "0";
                            }
                        }
                    }
                }
                matcher.appendReplacement(sb, Matcher.quoteReplacement(val));
            }
            matcher.appendTail(sb);

            String processed = sb.toString();

            // Evaluate arithmetic expression using internal evaluator
            return evaluateExpression(processed);
        } catch (StackOverflowError so) {
            so.printStackTrace();
            return Double.NaN;
        } catch (Exception e) {
            e.printStackTrace();
            return Double.NaN;
        }
    }

    // Find and replace functions (supports nested functions inside args)
    private String processAllFunctions(String expr) {
        // Keep applying replacements until no function patterns are left
        // functionName(args...) where args can contain commas, nested parentheses
        boolean changed = true;
        String working = expr;
        while (changed) {
            changed = false;
            // regex finds something like NAME(...) where NAME is letters
            Pattern p = Pattern.compile("([A-Z]+)\\(" , Pattern.CASE_INSENSITIVE);
            Matcher m = p.matcher(working);
            if (!m.find()) break;
            // We must scan manually to capture full parentheses groups
            StringBuilder sb = new StringBuilder();
            int i = 0;
            while (i < working.length()) {
                char c = working.charAt(i);
                if (Character.isLetter(c)) {
                    // possible start of function name
                    int startName = i;
                    while (i < working.length() && Character.isLetter(working.charAt(i))) i++;
                    if (i < working.length() && working.charAt(i) == '(') {
                        String fname = working.substring(startName, i).toUpperCase();
                        int parenStart = i;
                        int level = 0;
                        int j = i;
                        for (; j < working.length(); j++) {
                            char cc = working.charAt(j);
                            if (cc == '(') level++;
                            else if (cc == ')') {
                                level--;
                                if (level == 0) {
                                    break;
                                }
                            }
                        }
                        if (j >= working.length()) {
                            // mismatched parentheses — just append remainder and break
                            sb.append(working.substring(startName));
                            i = working.length();
                            break;
                        } else {
                            String inner = working.substring(parenStart + 1, j); // content inside parentheses
                            // compute function value
                            Double val = applyFunction(fname, inner.trim());
                            if (val == null || Double.isNaN(val)) {
                                sb.append("0"); // On error, substitute 0
                            } else {
                                // replace with numeric literal (unformatted)
                                sb.append(String.valueOf(val));
                            }
                            i = j + 1;
                            changed = true;
                            continue;
                        }
                    } else {
                        // not a function, append char at startName and continue
                        sb.append(working.charAt(startName));
                        i = startName + 1;
                    }
                } else {
                    sb.append(c);
                    i++;
                }
            }
            working = sb.toString();
        }
        return working;
    }

    // Apply a named function; returns null or NaN on error
    private Double applyFunction(String name, String argsText) {
        try {
            String fname = name.toUpperCase();
            List<String> args = splitArgs(argsText);
            List<Double> numbers = new ArrayList<>();

            // Collect numeric values from all args
            for (String a : args) {
                a = a.trim();
                if (a.isEmpty()) continue;

                if (a.contains(":")) {
                    numbers.addAll(getRangeValuesAll(a));
                } else {
                    int[] cell = parseCell(a);
                    if (cell != null) {
                        String raw = safeCell(cell[0], cell[1]);
                        if (raw != null && raw.startsWith("=")) {
                            double v = evaluateFormula(raw);
                            numbers.add(Double.isNaN(v) ? 0.0 : v);
                        } else {
                            numbers.add(parseDoubleOrZero(raw));
                        }
                    } else {
                        try {
                            numbers.add(Double.parseDouble(a));
                        } catch (NumberFormatException ex) {
                            // Evaluate expression with cell references
                            String expanded = a;
                            Pattern pattern = Pattern.compile("([A-Z]+)([0-9]+)");
                            Matcher matcher = pattern.matcher(expanded);
                            StringBuffer sb = new StringBuffer();
                            while (matcher.find()) {
                                String token = matcher.group();
                                int[] rc = parseCell(token);
                                String val = "0";
                                if (rc != null) {
                                    String raw = safeCell(rc[0], rc[1]);
                                    if (raw != null && raw.startsWith("=")) {
                                        double nested = evaluateFormula(raw);
                                        if (!Double.isNaN(nested)) val = String.valueOf(nested);
                                    } else {
                                        val = String.valueOf(parseDoubleOrZero(raw));
                                    }
                                }
                                matcher.appendReplacement(sb, Matcher.quoteReplacement(val));
                            }
                            matcher.appendTail(sb);
                            double eval = evaluateExpression(sb.toString());
                            numbers.add(Double.isNaN(eval) ? 0.0 : eval);
                        }
                    }
                }
            }

            // Precompute commonly used values for interrelated functions
            double sum = numbers.stream().mapToDouble(Double::doubleValue).sum();
            double avg = numbers.isEmpty() ? 0.0 : numbers.stream().mapToDouble(Double::doubleValue).average().orElse(0.0);
            double min = numbers.isEmpty() ? 0.0 : numbers.stream().mapToDouble(Double::doubleValue).min().orElse(0.0);
            double max = numbers.isEmpty() ? 0.0 : numbers.stream().mapToDouble(Double::doubleValue).max().orElse(0.0);
            double prod = 1.0;
            for (double d : numbers) prod *= d;

            switch (fname) {
                case "SUM": return sum;
                case "AVG":
                case "AVERAGE": return avg;
                case "MIN": return min;
                case "MAX": return max;
                case "COUNT": return (double) numbers.size();
                case "MEDIAN":
                    if (numbers.isEmpty()) return 0.0;
                    Collections.sort(numbers);
                    int n = numbers.size();
                    return (n % 2 == 1) ? numbers.get(n / 2) : (numbers.get(n / 2 - 1) + numbers.get(n / 2)) / 2.0;
                case "MODE":
                    if (numbers.isEmpty()) return 0.0;
                    Map<Double, Integer> freq = new HashMap<>();
                    for (Double d : numbers) freq.put(d, freq.getOrDefault(d, 0) + 1);
                    double mode = numbers.get(0);
                    int best = 0;
                    for (Map.Entry<Double, Integer> e : freq.entrySet()) {
                        if (e.getValue() > best || (e.getValue() == best && e.getKey() < mode)) {
                            mode = e.getKey();
                            best = e.getValue();
                        }
                    }
                    return mode;
                case "STDEV":
                    if (numbers.size() <= 1) return 0.0;
                    double sumsq = 0.0;
                    for (double d : numbers) sumsq += (d - avg) * (d - avg);
                    return Math.sqrt(sumsq / (numbers.size() - 1));
                case "RANGE": return max - min;
                case "PRODUCT": return numbers.isEmpty() ? 0.0 : prod;
                case "ABS": return numbers.isEmpty() ? 0.0 : Math.abs(numbers.get(0));
                case "SQRT": return numbers.isEmpty() ? 0.0 : Math.sqrt(numbers.get(0));
                case "MEAN":
                    return numbers.isEmpty() ? 0.0 : Math.pow(prod, 1.0 / numbers.size()); // geometric mean uses PRODUCT
                default: return Double.NaN;
            }

        } catch (Exception ex) {
            ex.printStackTrace();
            return Double.NaN;
        }
    }


    // Helper: split args at commas not inside nested parentheses
    private List<String> splitArgs(String s) {
        List<String> out = new ArrayList<>();
        if (s == null) return out;
        int level = 0;
        StringBuilder cur = new StringBuilder();
        for (int i = 0; i < s.length(); i++) {
            char c = s.charAt(i);
            if (c == ',' && level == 0) {
                out.add(cur.toString().trim());
                cur.setLength(0);
            } else {
                if (c == '(') level++;
                else if (c == ')') if (level > 0) level--;
                cur.append(c);
            }
        }
        if (cur.length() > 0) out.add(cur.toString().trim());
        return out;
    }

    // Expand ranges (A1:B2) and also single cell references into a list of numeric values (including nested formulas)
    private List<Double> getRangeValuesAll(String range) {
        List<Double> values = new ArrayList<>();
        if (range == null || range.isEmpty()) return values;
        String[] parts = range.split(":");
        if (parts.length == 1) {
            int[] cell = parseCell(parts[0].trim());
            if (cell != null && validCell(cell[0], cell[1])) {
                String raw = safeCell(cell[0], cell[1]);
                if (raw == null || raw.isEmpty()) {
                    values.add(0.0);
                } else if (raw.startsWith("=")) {
                    double res = evaluateFormula(raw);
                    values.add(Double.isNaN(res) ? 0.0 : res);
                } else {
                    values.add(parseDoubleOrZero(raw));
                }
            }
            return values;
        } else if (parts.length == 2) {
            int[] start = parseCell(parts[0].trim());
            int[] end = parseCell(parts[1].trim());
            if (start == null || end == null) return values;
            int minRow = Math.min(start[0], end[0]);
            int maxRow = Math.max(start[0], end[0]);
            int minCol = Math.min(start[1], end[1]);
            int maxCol = Math.max(start[1], end[1]);
            for (int r = minRow; r <= maxRow; r++) {
                for (int c = minCol; c <= maxCol; c++) {
                    if (validCell(r, c)) {
                        String raw = safeCell(r, c);
                        if (raw == null || raw.isEmpty()) {
                            values.add(0.0);
                        } else if (raw.startsWith("=")) {
                            double res = evaluateFormula(raw);
                            values.add(Double.isNaN(res) ? 0.0 : res);
                        } else {
                            values.add(parseDoubleOrZero(raw));
                        }
                    }
                }
            }
        }
        return values;
    }

    private boolean validCell(int r, int c) {
        return r >= 0 && r < rows && c >= 0 && c < cols && r < sheet.size() && c < sheet.get(r).size();
    }

    private String safeCell(int r, int c) {
        if (!validCell(r, c)) return "";
        return sheet.get(r).get(c);
    }

    private double parseDoubleOrZero(String s) {
        if (s == null) return 0.0;
        try {
            return Double.parseDouble(s);
        } catch (NumberFormatException ex) {
            return 0.0;
        }
    }

    private int[] parseCell(String ref) {
        try {
            ref = ref.trim().toUpperCase();
            Matcher m = Pattern.compile("([A-Z]+)([0-9]+)").matcher(ref);
            if (m.matches()) {
                String colPart = m.group(1);
                int col = 0;
                for (char c : colPart.toCharArray()) {
                    col = col * 26 + (c - 'A' + 1);
                }
                col--; // zero-index
                int rowIdx = Integer.parseInt(m.group(2)) - 1;
                return new int[]{rowIdx, col};
            }
        } catch (Exception ignored) {}
        return null;
    }

    /**
     * Evaluate arithmetic expression (numbers, + - * / ^, parentheses).
     * Uses shunting-yard algorithm to build RPN and then evaluate.
     */
    private double evaluateExpression(String expr) {
        try {
            // Tokenize
            List<String> tokens = tokenizeExpression(expr);
            if (tokens.isEmpty()) return Double.NaN;

            // Convert to RPN
            List<String> rpn = toRPN(tokens);
            if (rpn.isEmpty()) return Double.NaN;

            // Evaluate RPN
            Deque<Double> stack = new ArrayDeque<>();
            for (String tok : rpn) {
                if (isNumber(tok)) {
                    stack.push(Double.parseDouble(tok));
                } else {
                    if (tok.equals("+") || tok.equals("-") || tok.equals("*") || tok.equals("/") || tok.equals("^")) {
                        if (stack.size() < 2) return Double.NaN;
                        double b = stack.pop();
                        double a = stack.pop();
                        switch (tok) {
                            case "+": stack.push(a + b); break;
                            case "-": stack.push(a - b); break;
                            case "*": stack.push(a * b); break;
                            case "/":
                                if (b == 0) return Double.NaN;
                                stack.push(a / b); break;
                            case "^": stack.push(Math.pow(a, b)); break;
                        }
                    } else {
                        return Double.NaN; // unknown token
                    }
                }
            }
            if (stack.size() != 1) return Double.NaN;
            return stack.pop();
        } catch (Exception ex) {
            ex.printStackTrace();
            return Double.NaN;
        }
    }

    // Tokenizer: recognizes numbers, operators, parentheses. Handles unary minus by converting to (0 - x) pattern in tokens.
    private List<String> tokenizeExpression(String expr) {
        List<String> tokens = new ArrayList<>();
        if (expr == null) return tokens;
        String s = expr.trim();
        if (s.isEmpty()) return tokens;

        // remove spaces
        s = s.replaceAll("\\s+", "");

        int i = 0;
        while (i < s.length()) {
            char c = s.charAt(i);
            if ((c >= '0' && c <= '9') || c == '.') {
                int j = i + 1;
                while (j < s.length() && ((s.charAt(j) >= '0' && s.charAt(j) <= '9') || s.charAt(j) == '.')) j++;
                tokens.add(s.substring(i, j));
                i = j;
            } else if (c == '+' || c == '-' || c == '*' || c == '/' || c == '^') {
                // handle unary minus: if at start or after '(' or another operator, treat as unary
                if (c == '-') {
                    if (tokens.isEmpty() || tokens.get(tokens.size() - 1).equals("(") ||
                            tokens.get(tokens.size() - 1).equals("+") ||
                            tokens.get(tokens.size() - 1).equals("-") ||
                            tokens.get(tokens.size() - 1).equals("*") ||
                            tokens.get(tokens.size() - 1).equals("/") ||
                            tokens.get(tokens.size() - 1).equals("^")) {
                        // unary minus: convert to (0 - number) pattern by pushing "0" and "-" operator
                        tokens.add("0");
                        tokens.add("-");
                        i++;
                        continue;
                    }
                }
                tokens.add(String.valueOf(c));
                i++;
            } else if (c == '(' || c == ')') {
                tokens.add(String.valueOf(c));
                i++;
            } else {
                // unknown char: skip (or invalidate) - allow commas inside expressions to be ignored
                i++;
            }
        }
        return tokens;
    }

    private boolean isNumber(String s) {
        if (s == null) return false;
        try {
            Double.parseDouble(s);
            return true;
        } catch (NumberFormatException ex) {
            return false;
        }
    }

    // precedence and associativity
    private int precedence(String op) {
        switch (op) {
            case "+": case "-": return 1;
            case "*": case "/": return 2;
            case "^": return 3;
        }
        return 0;
    }

    private boolean isRightAssociative(String op) {
        return "^".equals(op);
    }

    // Convert to Reverse Polish Notation (shunting-yard)
    private List<String> toRPN(List<String> tokens) {
        List<String> output = new ArrayList<>();
        Deque<String> ops = new ArrayDeque<>();

        for (String tok : tokens) {
            if (isNumber(tok)) {
                output.add(tok);
            } else if (tok.equals("+") || tok.equals("-") || tok.equals("*") || tok.equals("/") || tok.equals("^")) {
                while (!ops.isEmpty() && !ops.peek().equals("(")) {
                    String top = ops.peek();
                    if ((isRightAssociative(tok) && precedence(tok) < precedence(top)) ||
                            (!isRightAssociative(tok) && precedence(tok) <= precedence(top))) {
                        output.add(ops.pop());
                    } else {
                        break;
                    }
                }
                ops.push(tok);
            } else if (tok.equals("(")) {
                ops.push(tok);
            } else if (tok.equals(")")) {
                while (!ops.isEmpty() && !ops.peek().equals("(")) {
                    output.add(ops.pop());
                }
                if (!ops.isEmpty() && ops.peek().equals("(")) {
                    ops.pop();
                } else {
                    return Collections.emptyList();
                }
            } else {
                return Collections.emptyList();
            }
        }

        while (!ops.isEmpty()) {
            String t = ops.pop();
            if (t.equals("(") || t.equals(")")) return Collections.emptyList(); // mismatched
            output.add(t);
        }

        return output;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            MiniExcel app = new MiniExcel();
            app.model.setRawValueAt("22",0,0);
            app.model.setRawValueAt("12",1,0);
            app.model.setRawValueAt("=SUM(A1:A2)",2,0);
            app.model.fireTableDataChanged();
        });
    }
}