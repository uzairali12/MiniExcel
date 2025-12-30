# 📊 MiniExcel – Java Swing Spreadsheet Application

## 📌 Overview

**MiniExcel** is a lightweight spreadsheet application developed using **Java Swing**. It simulates essential spreadsheet features such as cell formulas, row/column manipulation, undo/redo functionality, and CSV file handling.
The project focuses on applying **Data Structures and Algorithms (DSA)** concepts in a real-world, GUI-based application.

---

## ✨ Features

* Spreadsheet-like grid using `JTable`
* Insert and delete rows and columns
* Formula bar for cell expressions
* Supported formulas and functions:

  * `SUM`, `AVG`, `MEAN`, `MIN`, `MAX`, `COUNT`
  * `MEDIAN`, `MODE`, `STDEV`
  * `ABS`, `SQRT`, `PRODUCT`, `RANGE`
* Undo / Redo functionality
* Copy, Cut, and Paste operations
* Save and load spreadsheets in **CSV format**

---
## 🛠 Screenshots

<img width="975" height="794" alt="image" src="https://github.com/user-attachments/assets/8e865689-cbee-4116-9490-cdcf0ef38d2a" />
<img width="975" height="462" alt="image" src="https://github.com/user-attachments/assets/f86039c8-a9c8-40e2-ba69-a15250c45d8c" />
<img width="975" height="598" alt="image" src="https://github.com/user-attachments/assets/17c1a47f-b281-4454-8ea4-81f2c3690bce" />

---
## 🛠 Technologies Used

* **Language:** Java (JDK 8 or higher)
* **GUI:** Java Swing (`javax.swing`)
* **Data Structures:** List, Stack, HashMap, Deque
* **File Handling:** CSV using `java.io`
* **Formula Parsing:** Regular Expressions & Shunting-Yard Algorithm

---

## 📂 Project Structure

```
MiniExcel/
│
├── MiniExcel.java        # Main application file
├── README.md             # Project documentation
└── sample.csv            # Optional test CSV file
```

> ⚠️ The entire application is implemented in a **single Java file** (`MiniExcel.java`) for simplicity.

---

## ⬇️ Download & Setup Instructions

### Step 1: Download the Project

You can download the project in one of the following ways:

**Option 1: Download ZIP**

1. Click on **Code → Download ZIP**
2. Extract the ZIP file on your system

**Option 2: Clone Repository**

```bash
git clone https://github.com/your-username/MiniExcel.git
```

---

### Step 2: Install Java

Ensure **Java JDK 8 or higher** is installed.

Check installation:

```bash
java -version
```

If not installed, download from:

* [https://www.oracle.com/java/technologies/javase-downloads.html](https://www.oracle.com/java/technologies/javase-downloads.html)
  or
* [https://adoptium.net/](https://adoptium.net/)

---

### Step 3: Open in IDE (Recommended)

You may use any Java IDE:

* **Eclipse**
* **IntelliJ IDEA**
* **NetBeans**

Steps:

1. Open IDE
2. Select **Open Project** or **Import Project**
3. Choose the project folder
4. Open `MiniExcel.java`

---

### Step 4: Run the Application

1. Locate `MiniExcel.java`
2. Run the file:

   * **Right-click → Run**
   * or click the **Run ▶ button**

The MiniExcel window will launch.

---

## 🧑‍💻 How to Use MiniExcel

### Editing Cells

* Click any cell and type a value
* Use formulas starting with `=`
  Example:

  ```
  =SUM(A1:A5)
  ```

### Formula Bar

* Use the formula bar to enter or edit formulas
* Selected cell formula/value appears automatically

### Insert / Delete Rows & Columns

* Use the **menu bar** (Insert / Delete options)

### Copy, Cut, Paste

* Select cells
* Use menu options or toolbar buttons

### Undo / Redo

* Available via **Edit menu**
* Implemented using stack-based state tracking

### Save Spreadsheet

* `File → Save`
* Saves data in **CSV format**

### Load Spreadsheet

* `File → Open`
* Load previously saved CSV file

---

## 🧠 Implementation Details

### Data Storage

```java
List<List<String>> sheet;
```

* Dynamic 2D structure for storing raw values or formulas

### Undo / Redo

```java
Stack<State> undoStack;
Stack<State> redoStack;
```

* Each `State` stores a snapshot of the spreadsheet

### Formula Evaluation

* Cell references parsed using **Regex**
* Arithmetic expressions converted to **Reverse Polish Notation (RPN)**
* Evaluation done using **Deque**
* Supports nested formulas and range-based functions

### File Handling

* CSV files handled using `BufferedReader` and `PrintWriter`
* Proper handling of commas and special characters

---

## ⏱ Time Complexity Summary

| Operation            | Time Complexity                 |
| -------------------- | ------------------------------- |
| Cell access/update   | O(1)                            |
| Insert/Delete row    | O(n)                            |
| Insert/Delete column | O(m)                            |
| Formula evaluation   | O(n)                            |
| Range functions      | O(r × c)                        |
| Undo / Redo          | O(1) (snapshot: O(rows × cols)) |
| Save / Load CSV      | O(rows × cols)                  |

---

## ✅ Conclusion

MiniExcel demonstrates how **DSA concepts** such as stacks, lists, hash maps, and expression parsing algorithms can be effectively used in a **GUI-based desktop application**.
It bridges the gap between theoretical data structures and practical software development.

---

## 🚀 Future Improvements

* Cell formatting (colors, fonts, borders)
* Multiple sheet support
* Advanced formulas (IF, VLOOKUP)
* Performance optimization for large datasets
* Chart and graph generation

---

## 📚 References

* Java API Documentation
* Robert Lafore – *Data Structures and Algorithms in Java*
* Edsger W. Dijkstra – Shunting-Yard Algorithm
* RFC 4180 – CSV File Format Specification

---



