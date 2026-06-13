# 📁 ADODB – Advanced VBA Class for Text and Binary File Handling

![VBA](https://img.shields.io/badge/VBA-Excel-blue)
![License](https://img.shields.io/badge/License-MIT-green)

A complete and robust VBA class designed to read, write, and manipulate text or binary files using ADODB.Stream.  
It provides a simple, consistent, and secure interface for file operations, including native dialog boxes and detailed operation tracking.

---

## ⭐ Why use this class?

### ✔️ A modern alternative to native VBA file functions
VBA methods like `Open`, `Input`, `Line Input`, and `Print` are limited and unreliable with modern encodings.  
`ADODB.Stream` offers better stability, performance, and encoding support.

### ✔️ Automatic encoding management
UTF‑8, UTF‑16, ANSI, binary…  
The class hides all complexity and lets you choose the encoding easily.

### ✔️ Handles large files
Automatically detects large files (> 4 GB) and optimizes reading/writing to avoid memory issues.

### ✔️ Built‑in user interface
File selection, Save As, folder selection…  
No Windows API required.

### ✔️ Unified API
Same logic for text and binary files.  
No need to switch between different VBA syntaxes.

### ✔️ Safe and robust
- File and folder existence checks  
- Clean stream management  
- Byte and record counters  
- End‑of‑file detection  

### ✔️ Ideal for professional projects
Designed to be:
- reusable  
- stable  
- documented  
- easy to integrate  

---

## ✨ Main Features

### 🔧 File configuration
- FileType: text or binary  
- Encoding / EncodingTxt  
- LineSeparator  

### 🔒 Access management
- AccessType: read or write  
- FileName  
- Stream: direct access to ADODB.Stream  

### 🖥️ User interface
- SelectFile  
- SelectFileSaveAs  
- SelectFolder  

### 📊 Operation tracking
- RecordsRead / RecordsWritten  
- BytesRead / BytesWritten  

### 🧪 Utilities
- FileExists  
- FolderExists  
- IsLargeFile  
- FileLength  

---

---

## 📦 Project structure

```
ADODB/
 ├── ADODB.cls
 ├── AdoDB_Enum.bas
 ├── ExempleADODB.bas
 ├── LICENSE
 └── README.md
```

---

## 🛠️ Requirements

- Microsoft Excel / VBA  
- Microsoft ActiveX Data Objects x.x Library reference

---

## 📄 License

Distributed under the MIT License.

---

## 🤝 Contributing

Contributions are welcome:  
- suggestions  
- bug fixes  
- new features  

Open an issue or a pull request.
