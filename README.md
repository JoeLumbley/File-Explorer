# 📁 File Explorer

**File Explorer** is a simple, fast, and user‑friendly file management application designed to make navigating, organizing, and manipulating files intuitive for all users. It combines a clean graphical interface with a powerful built‑in **Command Line Interface (CLI)** for users who prefer keyboard‑driven workflows.




<img width="1266" height="713" alt="103" src="https://github.com/user-attachments/assets/72f96d2a-56b0-4001-a72a-792eb503ae25" />




## Features

### 🗂 Graphical Interface
- **Navigation**: Browse directories using a tree view and list view.
- **File Operations**: Perform essential actions such as:
  - Change Directory (`cd`)
  - Copy files and directories (`copy`)
  - Move files and directories (`move`)
  - Delete files and directories (`delete`)
- **Clipboard Support**: Cut, copy, and paste files and folders.
- **Context Menu**: Right‑click menus for quick access to common actions.
- **History Tracking**: Move backward and forward through previously visited directories.
- **File Type Icons**: Visual indicators for different file types.
- **Status Bar**: Real‑time feedback on operations and system status.

### 💻 Integrated Command Line Interface (CLI)
For users who enjoy the speed and precision of typed commands, the built‑in CLI provides:

- Fast directory navigation (`cd`)
- File and folder operations (`copy`, `move`, `delete`, `rename`)
- Text file creation (`text`, `txt`)
- Search and result cycling (`find`, `findnext`)
- Contextual path handling (type a path to open it)
- Quoting rules for paths with spaces
- Helpful usage messages and error feedback
- A built‑in `help` command with full documentation

The CLI and GUI work together seamlessly, giving users the freedom to choose the workflow that suits them best.


[ Table of Contents](#table-of-contents)





---
---
---
---




## Why I’m Creating File Explorer




I decided to build my own File Explorer because I wanted to understand, from the ground up, how a core part of every operating system actually works. We all use file managers every day, but it’s easy to forget how much is happening behind the scenes—navigation history, sorting, file type detection, context menus, clipboard operations, lazy‑loading folder trees, and so much more. Re‑creating these features myself has been a great way to explore system I/O, UI design, event handling, and performance considerations in a real, hands‑on way.

This project isn’t meant to replace the built‑in Windows Explorer. Instead, it’s a learning tool: a space where I can experiment, break things, fix them, and understand why they work the way they do.













## What I Hope Learners Get From This

This project is designed for anyone who wants to understand how real applications work, from beginners taking their first steps, to experienced developers exploring deeper architectural ideas. My hope is that you come away with:

### **A clearer understanding of how file systems are accessed and managed**  
By looking at the code behind navigation, file operations, and directory structures, you can see how your operating system performs these tasks under the hood.

### **Insight into building a real Windows Forms application**  
The project demonstrates UI layout, event‑driven programming, keyboard shortcuts, tooltips, context menus, and the small design decisions that make an interface feel intuitive and predictable.

### **Practical examples of organizing and structuring a larger project**  
You’ll find subsystems for navigation history, sorting logic, search functionality, and file-type mapping-all working together in a cohesive, maintainable way.

### **Confidence to modify, extend, or build your own tools**  
Everything is open-source under the MIT License, so you’re free to explore, customize, or reuse any part of the codebase in your own applications.

### **A reminder that even “simple” tools contain fascinating engineering challenges**  
Re-creating something familiar is one of the most effective ways to deepen your understanding. Much like how art students copy the masters to study technique, rebuilding a tool like File Explorer reveals the subtle decisions and hidden complexity behind everyday software.


If you’re curious, the GitHub repository includes the full source code and documentation. I’d love to hear your thoughts, suggestions, or ideas for future features. This project is as much about learning as it is about building something functional, and I’m excited to share that journey with you.












---
---
---
---
















## Table of Contents


- Code Walkthrough
  
  - [📦 MoveFileOrDirectory](#movefileordirectory)

  - [✏️ RenameFileOrDirectory](#renamefileordirectory)

  - [📋 CopyDirectory](#copydirectory)

  - [➡️ NavigateTo](#navigateto)



- [⌨️ Keyboard Shortcuts](#keyboard-shortcuts)




- [🖥️ Command Line Interface (CLI)](#command-line-interface)
  - [Features Overview](#-features-overview)
  - [Commands](#-commands)
    - [cd — Change Directory](#-change-directory--cd)
    - [mkdir / make — Create Directory](#-create-directory--mkdir-make)
    - [copy — Copy Files or Folders](#-copy--copy)
    - [move — Move Files or Folders](#-move--move)
    - [delete — Delete Files or Folders](#-delete--delete)
    - [rename — Rename Files or Folders](#-rename--rename)
    - [text / txt — Create Text Files](#-create-text-file--text-txt)
    - [open — Open Files or Folders](#-open--open)
    - [find / search — Search](#-search--find-search)
    - [findnext / searchnext — Next Search Result](#-next-search-result--findnext-searchnext)
    - [help — Show Help](#-help--help)
  - [Quoting Rules](#-quoting-rules-important)
  - [Contextual Navigation](#-contextual-navigation)
  - [Example Session](#-example-session)


- [📄 License](#license)

- [🧬 Clones](#clones)



---
---
---
---






























# Command Line Interface

The **Command Line Interface (CLI)** is an integrated text‑based command system inside the File Explorer application. It allows users to navigate folders, manage files, and perform common operations quickly using typed commands.





<img width="1266" height="713" alt="101" src="https://github.com/user-attachments/assets/11164542-dfc4-4885-93d0-e667a5630b4d" />



The CLI is designed to be:

- **Fast** — no menus, no dialogs  
- **Predictable** — clear rules and consistent behavior  
- **Beginner‑friendly** — helpful messages and examples  
- **Powerful** — supports navigation, search, file operations, and more  

[ Table of Contents](#table-of-contents)

---

## 🚀 Features Overview

### ✔ Navigation  
- Change directories  
- Open files directly  
- Navigate to folders by typing their path  
- Supports paths with spaces using quotes  

### ✔ File & Folder Operations  
- Create, copy, move, rename, and delete  
- Works with both files and directories  
- Handles quoted paths safely  

### ✔ Search  
- Search the current folder  
- Cycle through results with `findnext`  
- Highlights and selects results in the UI  

### ✔ Contextual Behavior  
If a command doesn’t match a known keyword, the CLI checks:

- **Is it a folder?** → Navigate to it  
- **Is it a file?** → Open it  
- Otherwise → “Unknown command”  

This makes the CLI feel natural and forgiving.

[ Table of Contents](#table-of-contents)

---

# 🧭 Commands

Below is the complete list of supported commands, including syntax, descriptions, and examples.

---

## 📂 Change Directory — `cd`

**Usage:**  
```
cd [directory]
```

**Description:**  
Changes the current working directory.

**Examples:**  
```
cd C:\
cd "C:\My Folder"
```

[ Table of Contents](#table-of-contents)

---

## 📁 Create Directory — `mkdir`, `make`

**Usage:**  
```
mkdir [directory_path]
```

**Description:**  
Creates a new folder.

**Examples:**  
```
mkdir C:\newfolder
make "C:\My New Folder"
```

[ Table of Contents](#table-of-contents)

---

## 📄 Copy — `copy`

**Usage:**  
```
copy [source] [destination]
```

**Description:**  
Copies a file or folder to a destination directory.

**Examples:**  
```
copy C:\folderA\file.txt C:\folderB
copy "C:\folder A" "C:\folder B"
```

[ Table of Contents](#table-of-contents)

---

## 📦 Move — `move`

**Usage:**  
```
move [source] [destination]
```

**Description:**  
Moves a file or folder to a new location.

**Examples:**  
```
move C:\folderA\file.txt C:\folderB\file.txt
move "C:\folder A\file.txt" "C:\folder B\renamed.txt"
```

[ Table of Contents](#table-of-contents)

---

## 🗑 Delete — `delete`

**Usage:**  
```
delete [file_or_directory]
```

**Description:**  
Deletes a file or folder.

**Examples:**  
```
delete C:\file.txt
delete "C:\My Folder"
```

[ Table of Contents](#table-of-contents)

---

## ✏ Rename — `rename`

**Usage:**  
```
rename [source_path] [new_name]
```

**Important:**  
Paths containing spaces **must** be enclosed in quotes.

**Examples:**  
```
rename "C:\folder\oldname.txt" "newname.txt"
rename "C:\folder\old name.txt" "new name.txt"
```

[ Table of Contents](#table-of-contents)

---

## 📝 Create Text File — `text`, `txt`

**Usage:**  
```
text [file_path]
```

**Description:**  
Creates a new text file at the specified path and opens it.

**Example:**  
```
text "C:\folder\example.txt"
```

If no file name is provided, the CLI creates a new file named:  
```
New Text File.txt
```

[ Table of Contents](#table-of-contents)

---

## 📂 Open — `open`

**Usage:**  
```
open [file_or_directory]
```

**Description:**  
Opens a file with its default application, or navigates into a folder.

If no path is provided, the command opens the **currently selected** file or folder in the File Explorer list.

**Examples:**  
```
open C:\folder\file.txt
open "C:\My Folder"
open
```

**Behavior:**

- If the target is a **file** → opens it using the default program  
- If the target is a **folder** → navigates into it  
- If nothing is selected and no path is provided → shows usage help  
- Supports quoted paths with spaces  

[ Table of Contents](#table-of-contents)

---


## 🔍 Search — `find`, `search`

**Usage:**  
```
find [search_term]
```

**Description:**  
Searches the current folder for files or folders containing the term.

**Example:**  
```
find report
```

If results are found:

- The first result is automatically selected  
- The status bar shows how many matches were found  

[ Table of Contents](#table-of-contents)

---

## ⏭ Next Search Result — `findnext`, `searchnext`

**Usage:**  
```
findnext
```

**Description:**  
Cycles to the next result from the previous search.  
Wraps around when reaching the end.

[ Table of Contents](#table-of-contents)

---

## ❌ Exit — `exit`, `quit`

Closes the application.

---

## ❓ Help — `help`

Displays the full list of commands.

---

# 🧠 Quoting Rules (Important)

Paths containing spaces **must** be enclosed in quotes:

```
"C:\My Folder"
"C:\Users\Joe\My File.txt"
```

This applies to:

- `cd`
- `copy`
- `move`
- `rename`
- `delete`
- `text`

The CLI will warn the user when quotes are required.

[ Table of Contents](#table-of-contents)

---

# 🧭 Contextual Navigation

If the user enters something that is **not** a command:

- If it’s a **folder path**, the CLI navigates to it  
- If it’s a **file path**, the CLI opens it  
- Otherwise, the CLI shows an “Unknown command” message  

This makes the CLI feel natural and forgiving.

[ Table of Contents](#table-of-contents)

---

# 🖥 Example Session

```
> cd C:\Users\Joe
Navigated To: C:\Users\Joe

> copy "C:\Users\Joe\file.txt" "C:\Users\Joe\Documents"
Copied file: file.txt to: C:\Users\Joe\Documents

> find report
Found 3 result(s). Showing result 1. Type findnext to move to the next match.

> findnext
Showing result 2 of 3

> help
(Displays full help text)
```

---

# 🎯 Summary

The File Explorer CLI provides:

- Fast directory navigation  
- Powerful file operations  
- Search with result cycling  
- Intelligent path handling  
- Clear feedback and usage messages  
- Beginner‑friendly quoting rules  
- Contextual file/folder opening  

It’s a flexible, efficient alternative to the graphical interface — perfect for users who enjoy command‑driven workflows.


[ Table of Contents](#table-of-contents)








---
---
---
---







# `RenameFileOrDirectory`

This method renames either a **file** or a **folder**, following a clear set of safety rules. Each rule protects the user from mistakes and helps them understand what’s happening.



<img width="1275" height="662" alt="040" src="https://github.com/user-attachments/assets/db463d52-0d3f-40c7-bb34-867deb1925d7" />




Below is the full code, then we’ll walk through it one step at a time.


```vb.net

    Private Sub RenameFileOrDirectory(sourcePath As String, newName As String)

        Dim newPath As String = Path.Combine(Path.GetDirectoryName(sourcePath), newName)

        ' Rule 1: Path must be absolute (start with C:\ or similar).
        ' Reject relative paths outright
        If Not Path.IsPathRooted(sourcePath) Then

            ShowStatus(IconDialog & " Rename failed: Path must be absolute. Example: C:\folder")

            Exit Sub

        End If

        ' Rule 2: Protected paths are never renamed.
        ' Check if the path is in the protected list
        If IsProtectedPathOrFolder(sourcePath) Then
            ' The path is protected; prevent rename

            ' Show user the directory so they can see it wasn't renamed.
            NavigateTo(sourcePath)

            ' Notify the user of the prevention so the user knows why it didn't rename.
            ShowStatus(IconProtect & "  Rename prevented for protected path or folder: " & sourcePath)

            Exit Sub

        End If

        Try

            ' Rule 3: If it’s a folder, rename the folder and show the new folder.
            ' If source is a directory
            If Directory.Exists(sourcePath) Then

                ' Rename directory
                Directory.Move(sourcePath, newPath)

                ' Navigate to the renamed directory
                NavigateTo(newPath)

                ShowStatus(IconSuccess & " Renamed Folder to: " & newName)

                ' Rule 4: If it’s a file, rename the file and show its folder.
                ' If source is a file
            ElseIf File.Exists(sourcePath) Then

                ' Rename file
                File.Move(sourcePath, newPath)

                ' Navigate to the directory of the renamed file
                NavigateTo(Path.GetDirectoryName(sourcePath))

                ShowStatus(IconSuccess & " Renamed File to: " & newName)

                ' Rule 5: If nothing exists at that path, explain the quoting rule for spaces.
            Else
                ' Path does not exist
                ShowStatus(IconError & " Renamed failed: No path. Paths with spaces must be enclosed in quotes. Example: rename ""[source_path]"" ""[new_name]"" e.g., rename ""C:\folder\old name.txt"" ""new name.txt""")

            End If

        Catch ex As Exception
            ' Rule 6: If anything goes wrong,show a status message.
            ShowStatus(IconError & " Rename failed: " & ex.Message)
            Debug.WriteLine("RenameFileOrDirectory Error: " & ex.Message)
        End Try

    End Sub

```




## 🔧 Method Signature

```vb.net
Private Sub RenameFileOrDirectory(sourcePath As String, newName As String)
```

- **sourcePath** — the full path to the file or folder you want to rename  
- **newName** — just the new name (not a full path)



## 🧱 Step 1 — Build the new full path

```vb.net
Dim newPath As String = Path.Combine(Path.GetDirectoryName(sourcePath), newName)
```

- Extracts the folder that contains the item  
- Combines it with the new name  
- Example:  
  - `sourcePath = "C:\Stuff\Old.txt"`  
  - `newName = "New.txt"`  
  - `newPath = "C:\Stuff\New.txt"`



## 🛑 Rule 1 — Path must be absolute

```vb.net
If Not Path.IsPathRooted(sourcePath) Then
    ShowStatus(IconDialog & " Rename failed: Path must be absolute. Example: C:\folder")
    Exit Sub
End If
```

Beginners often type relative paths like `folder\file.txt`.  
This rule stops the rename and explains the correct format.



## 🔒 Rule 2 — Protected paths are never renamed

```vb.net
If IsProtectedPathOrFolder(sourcePath) Then
    NavigateTo(sourcePath)
    ShowStatus(IconProtect & "  Rename prevented for protected path or folder: " & sourcePath)
    Exit Sub
End If
```

Some paths should never be renamed (system folders, app folders, etc.).  
This rule:

- Prevents the rename  
- Shows the user the original folder  
- Explains why the rename was blocked  

This is excellent for learner clarity.



## 🧪 Try/Catch — Safe execution zone

```vb.net
Try
    ...
Catch ex As Exception
    ShowStatus(IconError & " Rename failed: " & ex.Message)
    Debug.WriteLine("RenameFileOrDirectory Error: " & ex.Message)
End Try
```

Anything inside the `Try` block that fails will be caught and explained.  
Beginners get a friendly message instead of a crash.



## 📁 Rule 3 — If it’s a folder, rename the folder

```vb.net
If Directory.Exists(sourcePath) Then
    Directory.Move(sourcePath, newPath)
    NavigateTo(newPath)
    ShowStatus(IconSuccess & " Renamed Folder to: " & newName)
```

- Checks if the path points to a **directory**  
- Renames it  
- Shows the user the newly renamed folder  

This reinforces the idea that folders are “containers” and have their own identity.



## 📄 Rule 4 — If it’s a file, rename the file

```vb.net
ElseIf File.Exists(sourcePath) Then
    File.Move(sourcePath, newPath)
    NavigateTo(Path.GetDirectoryName(sourcePath))
    ShowStatus(IconSuccess & " Renamed File to: " & newName)
```

- Checks if the path points to a **file**  
- Renames it  
- Shows the user the folder containing the renamed file  

This keeps the UI consistent and predictable.



## ❓ Rule 5 — If nothing exists at that path, explain quoting rules

```vb.net
Else
    ShowStatus(IconError & " Renamed failed: No path. Paths with spaces must be enclosed in quotes. Example: rename ""[source_path]"" ""[new_name]"" e.g., rename ""C:\folder\old name.txt"" ""new name.txt""")
End If
```

Beginners often forget to quote paths with spaces.  
This rule:

- Detects the missing path  
- Explains the quoting rule  
- Gives both a template and a real example  

This is *excellent* pedagogy.



## ⚠️ Rule 6 — If anything goes wrong, show a clear error

```vb.net
Catch ex As Exception
    ShowStatus(IconError & " Rename failed: " & ex.Message)
    Debug.WriteLine("RenameFileOrDirectory Error: " & ex.Message)
End Try
```

- No crashes  
- Clear feedback  
- Debug info for you  



This method teaches six important rules:

1. **Paths must be absolute**  
2. **Protected paths cannot be renamed**  
3. **Folders are renamed and then shown**  
4. **Files are renamed and their folder is shown**  
5. **Missing paths trigger a helpful quoting explanation**  
6. **Any error is caught and explained safely**









<img width="1275" height="662" alt="041" src="https://github.com/user-attachments/assets/15de34d4-f14c-4968-a10d-2203af0c7130" />



[ Table of Contents](#table-of-contents)

---
---
---
---






# `CopyDirectory`

This method copies an entire directory — including all files and all subfolders — into a new destination. It uses **recursion**, meaning the method calls itself to handle deeper levels of folders.




<img width="1266" height="713" alt="086" src="https://github.com/user-attachments/assets/b40cb697-3ba8-4363-8e4b-1e7277661b1b" />



Below is the full code, then we’ll walk through it one step at a time.


```vb.net

    Private Async Function CopyDirectory(sourceDir As String, destDir As String) As Task
        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then
            ShowStatus(IconError & " Source directory not found: " & sourceDir)
            Return
        End If

        Try
            ShowStatus(IconCopy & " Creating destination directory: " & destDir)

            ' Create destination directory
            Try
                Directory.CreateDirectory(destDir)
            Catch ex As Exception
                ShowStatus(IconError & " Failed to create destination directory: " & ex.Message)
                Return
            End Try

            ShowStatus(IconCopy & " Copying files to destination directory: " & destDir)

            ' Copy files asynchronously
            For Each file In dirInfo.GetFiles()
                Try
                    Dim targetFilePath = Path.Combine(destDir, file.Name)
                    Await Task.Run(Sub() file.CopyTo(targetFilePath, overwrite:=True))
                    Debug.WriteLine("Copied file: " & targetFilePath) ' Log successful copy
                Catch ex As UnauthorizedAccessException
                    Debug.WriteLine("CopyDirectory Error (Unauthorized): " & ex.Message)
                    ShowStatus(IconError & " Unauthorized access: " & file.FullName)
                Catch ex As Exception
                    Debug.WriteLine("CopyDirectory Error: " & ex.Message)
                    ShowStatus(IconError & " Copy failed for file: " & file.FullName & " - " & ex.Message)
                End Try
            Next

            ShowStatus(IconCopy & " Copying subdirectories.")

            ' Copy subdirectories recursively asynchronously
            For Each subDir In dirInfo.GetDirectories()
                Dim newDest = Path.Combine(destDir, subDir.Name)
                Try
                    Await CopyDirectory(subDir.FullName, newDest)
                Catch ex As Exception
                    Debug.WriteLine("CopyDirectory Error: " & ex.Message)
                End Try
            Next

            ' Refresh the view to show the copied directory
            NavigateTo(destDir)

            ShowStatus(IconSuccess & " Copied into " & destDir)

        Catch ex As Exception
            ShowStatus(IconError & " Copy failed: " & ex.Message)
            Debug.WriteLine("CopyDirectory Error: " & ex.Message)
        End Try
    End Function

```


Here's a detailed breakdown of the updated `CopyDirectory` method , which now supports asynchronous file copying. This ensures that the UI remains responsive during the operation, especially when dealing with large directories.

## Updated Method Definition

```vb.net
Private Async Function CopyDirectory(sourceDir As String, destDir As String) As Task
```

- **sourceDir**: The folder you want to copy.
- **destDir**: The location where the copy should be created.

## Create a DirectoryInfo Object for the Source

```vb.net
Dim dirInfo As New DirectoryInfo(sourceDir)
```

- This creates a `DirectoryInfo` object that provides access to the folder's files and subfolders.

## Ensure the Source Directory Exists

```vb.net
If Not dirInfo.Exists Then
    ShowStatus(IconError & " Source directory not found: " & sourceDir)
    Return
End If
```

- Checks if the source directory exists. If not, it shows an error message and exits the method.

## Start a Try/Catch Block

```vb.net
Try
```

- Initiates a block to handle exceptions that may occur during the operation.

## Create the Destination Directory

```vb.net
Try
    Directory.CreateDirectory(destDir)
Catch ex As Exception
    ShowStatus(IconError & " Failed to create destination directory: " & ex.Message)
    Return
End Try
```

- Attempts to create the destination directory. If it fails, an error message is shown, and the method exits.

## Copy Files Asynchronously

```vb.net
For Each file In dirInfo.GetFiles()
    Try
        Dim targetFilePath = Path.Combine(destDir, file.Name)
        Await Task.Run(Sub() file.CopyTo(targetFilePath, overwrite:=True))
        Debug.WriteLine("Copied file: " & targetFilePath) ' Log successful copy
    Catch ex As UnauthorizedAccessException
        Debug.WriteLine("CopyDirectory Error (Unauthorized): " & ex.Message)
        ShowStatus(IconError & " Unauthorized access: " & file.FullName)
    Catch ex As Exception
        Debug.WriteLine("CopyDirectory Error: " & ex.Message)
        ShowStatus(IconError & " Copy failed for file: " & file.FullName & " - " & ex.Message)
    End Try
Next
```

- Iterates through each file in the source directory and copies it to the destination asynchronously. 
- Uses `Await Task.Run(...)` to ensure the UI remains responsive during the file copy process.
- Handles exceptions specifically for unauthorized access and general errors.

## Copy Subdirectories Recursively

```vb.net
For Each subDir In dirInfo.GetDirectories()
    Dim newDest = Path.Combine(destDir, subDir.Name)
    Try
        Await CopyDirectory(subDir.FullName, newDest)
    Catch ex As Exception
        Debug.WriteLine("CopyDirectory Error: " & ex.Message)
    End Try
Next
```

- For each subdirectory, constructs a new destination path and calls `CopyDirectory` recursively.
- Utilizes `Await` for asynchronous execution.

## Refresh the UI

```vb.net
NavigateTo(destDir)
```

- Updates the UI to reflect the new state after the copy operation is complete.

## Show Success Message

```vb.net
ShowStatus(IconSuccess & " Copied into " & destDir)
```

- Displays a success message once the operation is completed.

## Handle Any Errors

```vb.net
Catch ex As Exception
    ShowStatus(IconError & " Copy failed: " & ex.Message)
    Debug.WriteLine("CopyDirectory Error: " & ex.Message)
End Try
```

- If any errors occur during the overall operation, they are caught here, and an appropriate message is displayed.

## Summary

- **Asynchronous Support**: The method now uses `Async` and `Await` to ensure the application remains responsive during file operations.
- **Enhanced Error Handling**: Specific handling for unauthorized access and general exceptions.
- **User Feedback**: Continuous feedback is provided to the user throughout the process.

This updated method demonstrates how to effectively manage file copying operations while maintaining a responsive user interface.




<img width="1266" height="713" alt="087" src="https://github.com/user-attachments/assets/a2106467-9423-4698-b605-88a6e484a6ce" />




[ Table of Contents](#table-of-contents)

---











# 🌳 **Recursion Flow Diagram for `CopyDirectory`**

This diagram shows exactly how the `CopyDirectory` routine walks the tree.

Imagine your folder structure looks like this:

```
SourceDir
├── FileA.txt
├── FileB.txt
├── Sub1
│   ├── FileC.txt
│   └── Sub1A
│       └── FileD.txt
└── Sub2
    └── FileE.txt
```

The method processes it in this exact order.



# **High‑Level Recursion Flow**

```
CopyDirectory(SourceDir, DestDir)
│
├── Copy files in SourceDir
│
├── For each subdirectory:
│     ├── CopyDirectory(Sub1, DestDir/Sub1)
│     │     ├── Copy files in Sub1
│     │     ├── CopyDirectory(Sub1A, DestDir/Sub1/Sub1A)
│     │     │     ├── Copy files in Sub1A
│     │     │     └── (Sub1A has no more subfolders → return)
│     │     └── (Sub1 done → return)
│     │
│     └── CopyDirectory(Sub2, DestDir/Sub2)
│           ├── Copy files in Sub2
│           └── (Sub2 has no more subfolders → return)
│
└── All subdirectories processed → return to caller
```



# **Step‑By‑Step Call Stack Visualization**

This shows how the *call stack* grows and shrinks as recursion happens.

```
Call 1: CopyDirectory(SourceDir)
    ├── copies files
    ├── enters Sub1 → Call 2

Call 2: CopyDirectory(Sub1)
    ├── copies files
    ├── enters Sub1A → Call 3

Call 3: CopyDirectory(Sub1A)
    ├── copies files
    └── no subfolders → return to Call 2

Back to Call 2:
    └── Sub1 done → return to Call 1

Back to Call 1:
    ├── enters Sub2 → Call 4

Call 4: CopyDirectory(Sub2)
    ├── copies files
    └── no subfolders → return to Call 1

Back to Call 1:
    └── all done → return to caller
```



# **Indented Tree Showing Recursion Depth**

Each level of indentation = one level deeper in recursion.

```
CopyDirectory(SourceDir)
    CopyDirectory(Sub1)
        CopyDirectory(Sub1A)
    CopyDirectory(Sub2)
```

This is the simplest way to show the “shape” of recursion.



# **Narrative Version**

> 1. Start at the root folder.  
> 2. Copy its files.  
> 3. For each subfolder:  
>    - Step into it  
>    - Treat it like a brand‑new root  
>    - Copy its files  
>    - Repeat the process for its subfolders  
> 4. When a folder has no subfolders, return to the previous level.  
> 5. Continue until you climb all the way back to the top.











[ Table of Contents](#table-of-contents)

---
---
---
---








# MoveFileOrDirectory

<img width="1266" height="662" alt="073" src="https://github.com/user-attachments/assets/fe3323d2-0087-4ab8-8d00-755558bf2ccf" />

This walkthrough explains how the `MoveFileOrDirectory` routine works inside the file‑manager project. The goal of this method is to safely move files or directories while providing clear, user‑friendly feedback and preventing dangerous or confusing operations.

---

## 🧭 Overview

`MoveFileOrDirectory(source, destination)` performs a safe move operation with:

- **Input validation**  
- **Protected‑path checks**  
- **Self‑move and recursive‑move prevention**  
- **Automatic creation of destination directories**  
- **User‑visible navigation before and after the move**  
- **Clear status messages for every outcome**

This mirrors the project’s design philosophy:  
**Show → Confirm → Act → Show Result**

---

## 🧩 Full Method

```vb.net
Private Sub MoveFileOrDirectory(source As String, destination As String)
    Try
        ' Validate parameters
        If String.IsNullOrWhiteSpace(source) OrElse String.IsNullOrWhiteSpace(destination) Then
            ShowStatus(IconWarning & " Source or destination path is invalid.")
            Return
        End If

        ' If source and destination are the same, do nothing
        If String.Equals(source.TrimEnd("\"c), destination.TrimEnd("\"c), StringComparison.OrdinalIgnoreCase) Then
            ShowStatus(IconWarning & " Source and destination paths are the same. Move operation canceled.")
            Return
        End If

        ' Is source on the protected paths list?
        If IsProtectedPathOrFolder(source) Then
            ShowStatus(IconProtect & " Move operation prevented for protected path: " & source)
            Return
        End If

        ' Is destination on the protected paths list?
        If IsProtectedPathOrFolder(destination) Then
            ShowStatus(IconProtect & " Move operation prevented for protected path: " & destination)
            Return
        End If

        ' Prevent moving a directory into itself or its subdirectory
        If Directory.Exists(source) AndAlso
           (String.Equals(source.TrimEnd("\"c), destination.TrimEnd("\"c), StringComparison.OrdinalIgnoreCase) OrElse
            destination.StartsWith(source.TrimEnd("\"c) & "\", StringComparison.OrdinalIgnoreCase)) Then
            ShowStatus(IconWarning & " Cannot move a directory into itself or its subdirectory.")
            Return
        End If

        ' Check if the source is a file
        If File.Exists(source) Then

            ' Check if the destination file already exists
            If Not File.Exists(destination) Then

                ' Navigate to the directory of the source file
                NavigateTo(Path.GetDirectoryName(source))

                ShowStatus(IconDialog & "  Moving file to: " & destination)

                ' Ensure destination directory exists
                Directory.CreateDirectory(Path.GetDirectoryName(destination))

                File.Move(source, destination)

                ' Navigate to the destination folder
                NavigateTo(Path.GetDirectoryName(destination))

                ShowStatus(IconSuccess & "  Moved file to: " & destination)

            Else
                ShowStatus(IconWarning & " Destination file already exists.")
            End If

        ElseIf Directory.Exists(source) Then

            ' Check if the destination directory already exists
            If Not Directory.Exists(destination) Then

                ' Navigate to the directory being moved so the user can see it
                NavigateTo(source)

                ShowStatus(IconDialog & "  Moving directory to: " & destination)

                ' Ensure destination parent exists
                Directory.CreateDirectory(Path.GetDirectoryName(destination))

                ' Perform the move
                Directory.Move(source, destination)

                ' Navigate to the new location FIRST
                NavigateTo(destination)

                ' Now refresh the tree roots
                InitTreeRoots()

                ShowStatus(IconSuccess & "  Moved directory to: " & destination)

            Else
                ShowStatus(IconWarning & " Destination directory already exists.")
            End If

        Else
            ShowStatus(IconWarning & "  Move failed: Source path not found. Paths with spaces must be enclosed in quotes. Example: move ""C:\folder A"" ""C:\folder B""")
        End If

    Catch ex As Exception
        ShowStatus(IconError & " Move failed: " & ex.Message)
        Debug.WriteLine("MoveFileOrDirectory Error: " & ex.Message)
    End Try
End Sub
```

---

## 🧠 How It Works (Step‑By‑Step)

### 1. **Input Validation**
Rejects empty or whitespace paths to prevent confusing errors.

### 2. **Same‑Path Check**
If the source and destination resolve to the same location, the move is canceled.

### 3. **Protected Path Safety**
Both source and destination are checked against a protected‑paths list.  
Protected paths cannot be moved.

### 4. **Recursive Move Prevention**
The method prevents moving a directory into:

- Itself  
- One of its own subdirectories  

This avoids catastrophic recursive behavior.

### 5. **File Move Logic**
If the source is a file:

- Navigate to the file’s directory  
- Ensure the destination directory exists  
- Move the file  
- Navigate to the destination folder  
- Show success  

### 6. **Directory Move Logic**
If the source is a directory:

- Navigate into the directory being moved  
- Ensure the destination parent exists  
- Move the directory  
- Navigate to the new location  
- Refresh the tree  
- Show success  

### 7. **Error Handling**
Any exception results in:

- A user‑friendly status message  
- A debug log entry  

---

## 🎯 Design Philosophy

This method is built around **clarity, safety, and emotional transparency**:

- The user always sees *what is about to happen*  
- The user always sees *the result*  
- Dangerous operations are prevented  
- All actions are narrated through status messages  
- The UI updates in a way that reinforces the mental model  

This makes the file manager not just functional — but **guiding**.

<img width="1266" height="662" alt="074" src="https://github.com/user-attachments/assets/ea1ecd50-d794-40c9-9723-824487dfff1e" />









[ Table of Contents](#table-of-contents)

---
---
---
---






# `NavigateTo`

The `NavigateTo` method is responsible for navigating to a specified folder path.





Below is the full code, then we’ll walk through it one step at a time.


```vb.net


    Private Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
        ' Navigate to the specified folder path.
        ' Updates the current folder, path textbox, and file list.

        If String.IsNullOrWhiteSpace(path) Then Exit Sub

        ' Validate that the folder exists
        If Not Directory.Exists(path) Then
            MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ShowStatus(IconWarning & " Folder not found: " & path)
            Exit Sub
        End If

        ' If this method is called from a background thread, invoke it on the UI thread
        If txtPath.InvokeRequired Then
            txtPath.Invoke(New Action(Of String)(AddressOf NavigateTo), path, recordHistory)
            Return
        End If

        ShowStatus(IconNavigate & " Navigated To: " & path)

        currentFolder = path
        txtPath.Text = path
        PopulateFiles(path)

        If recordHistory Then
            ' Trim forward history if we branch
            If _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1 Then
                _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1))
            End If
            _history.Add(path)
            _historyIndex = _history.Count - 1
            UpdateNavButtons()
        End If

        UpdateFileButtons()
        UpdateEditButtons()
        UpdateEditContextMenu()
    End Sub

```





Here's a detailed breakdown of the `NavigateTo` method in VB.NET, which is responsible for navigating to a specified folder path. This method updates the UI components accordingly and manages navigation history.

## Method Overview

```vb.net
Private Sub NavigateTo(path As String, Optional recordHistory As Boolean = True)
```

- **path**: The folder path to navigate to.
- **recordHistory**: An optional boolean parameter that indicates whether the navigation should be recorded in history (default is `True`).

## Early Exit for Invalid Path

```vb.net
If String.IsNullOrWhiteSpace(path) Then Exit Sub
```

- Checks if the provided path is null, empty, or consists only of whitespace. If so, the method exits early.

## Validate Folder Existence

```vb.net
If Not Directory.Exists(path) Then
    MessageBox.Show("Folder not found: " & path, "Navigation", MessageBoxButtons.OK, MessageBoxIcon.Information)
    ShowStatus(IconWarning & " Folder not found: " & path)
    Exit Sub
End If
```

- Validates that the specified folder exists. If it does not, a message box is displayed to inform the user, and a warning status is shown. The method then exits.

## Handle UI Thread Invocation

```vb.net
If txtPath.InvokeRequired Then
    txtPath.Invoke(New Action(Of String)(AddressOf NavigateTo), path, recordHistory)
    Return
End If
```

- Checks if the method is being called from a background thread (indicated by `InvokeRequired`). If so, it invokes the method on the UI thread using `Invoke`, ensuring that UI updates are performed on the correct thread.

## Update UI Components

```vb.net
ShowStatus(IconNavigate & " Navigated To: " & path)

currentFolder = path
txtPath.Text = path
PopulateFiles(path)
```

- Updates the status to indicate the navigation action.
- Sets the `currentFolder` variable to the new path.
- Updates the text box (`txtPath`) with the new path.
- Calls `PopulateFiles(path)` to refresh the file list displayed in the UI.

## Record Navigation History

```vb.net
If recordHistory Then
    ' Trim forward history if we branch
    If _historyIndex >= 0 AndAlso _historyIndex < _history.Count - 1 Then
        _history.RemoveRange(_historyIndex + 1, _history.Count - (_historyIndex + 1))
    End If
    _history.Add(path)
    _historyIndex = _history.Count - 1
    UpdateNavButtons()
End If
```

- If `recordHistory` is `True`, it manages the navigation history:
  - Trims the forward history if the current index is not at the end of the history list.
  - Adds the new path to the history list.
  - Updates the `_historyIndex` to point to the last entry in the history.
  - Calls `UpdateNavButtons()` to refresh the navigation buttons (e.g., back and forward).

## Update Other UI Elements

```vb.net
UpdateFileButtons()
UpdateEditButtons()
UpdateEditContextMenu()
```

- Calls methods to update various UI elements related to file actions, editing, and context menus, ensuring they reflect the current state after navigation.

## Summary

The `NavigateTo` method effectively handles folder navigation with the following key features:

- **Input Validation**: Checks for valid paths and existence of directories.
- **Thread Safety**: Ensures UI updates occur on the correct thread.
- **User Feedback**: Provides immediate feedback on navigation actions and errors.
- **History Management**: Maintains a history of navigated folders, allowing for backward and forward navigation.
- **UI Updates**: Refreshes relevant UI components to reflect the current folder state.

This method is essential for any file management application, ensuring smooth and intuitive navigation for users.















---
---
---
---




---

# Keyboard Shortcuts

The File Explorer supports a set of convenient keyboard shortcuts to speed up navigation and file operations. These shortcuts mirror familiar behaviors from traditional file managers, making the interface fast and intuitive.

---

## 🧭 Navigation

| Shortcut | Action |
|---------|--------|
| **Alt + ←** (or **Backspace**) | Go back to the previous folder |
| **Alt + →** | Go forward to the next folder |
| **Alt + ↑** | Move up one level (parent directory) |
| **Ctrl + L** (or **Alt + D**, **F4**) | Focus/select the address bar |
| **F11** | Toggle full‑screen mode |

---

---

## 📁 File & Folder Operations

| Shortcut | Action |
|---------|--------|
| **F2** | Rename the selected file or folder |
| **Ctrl + Shift + N** | Create a new folder |
| **Ctrl + Shift + T** | Create a new text file |
| **Ctrl + O** | Open the selected file or folder or start an `open` command |
| **Ctrl + C** | Copy selected items |
| **Ctrl + V** | Paste items |
| **Ctrl + X** | Cut selected items |
| **Ctrl + A** | Select all items |
| **Ctrl + Shift + E** | Expand the selected folder or drive |
| **Ctrl + Shift + C** | Collapse the selected folder or drive |
| **Ctrl + D** (or **Delete**) | Delete selected item to the Recycle Bin |
---

## 🔍 Searching & Viewing

| Shortcut | Action |
|---------|--------|
| **Ctrl + F** | Start a search in the current folder |
| **F3**  | Select the next search result. |
| **F5** | Refresh the current folder view |

---
























[ Table of Contents](#table-of-contents)

---
---
---
---







## License
This project is licensed under the **MIT License**. You are free to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the software, with the inclusion of the copyright notice and permission notice in all copies or substantial portions of the software.










[ Table of Contents](#table-of-contents)

---
---
---
---














## Installation
To install and run the File Explorer application:
1. Clone the repository:
   ```bash
   git clone https://github.com/JoeLumbley/File-Explorer.git
   ```
2. Open the solution in Visual Studio.
3. Build the project and run it.














[ Table of Contents](#table-of-contents)

---
---
---
---










## Usage
- Launch the application to access your file system.
- Use the tree view on the left to navigate through folders.
- The list view on the right displays the contents of the selected directory.
- Use the text box to enter commands or navigate directly to a path.
- Right-click on files or folders for additional options.









[ Table of Contents](#table-of-contents)

---
---
---
---










## Acknowledgements
This project is inspired by traditional file explorers and aims to provide a simplified experience for managing files on Windows systems.

For more details, check the source code and documentation within the repository.















[ Table of Contents](#table-of-contents)

---
---
---
---












# Clones










<img width="1920" height="1080" alt="094" src="https://github.com/user-attachments/assets/0d9b3b5f-32c2-496c-a025-b2ec26e5276f" />












<img width="1920" height="1080" alt="088" src="https://github.com/user-attachments/assets/70f51b28-dcda-4803-9d94-2c61eeaa6d7b" />







<img width="1920" height="1080" alt="083" src="https://github.com/user-attachments/assets/763aac35-b8b8-410d-89a5-a8fd892d7219" />






<img width="1920" height="1080" alt="080" src="https://github.com/user-attachments/assets/680cab6e-3f3d-48bc-8075-9c1c830754ec" />





<img width="1920" height="1080" alt="079" src="https://github.com/user-attachments/assets/fdd74f0f-7cdd-452f-afa7-1bec1b118db2" />







<img width="1920" height="1080" alt="078" src="https://github.com/user-attachments/assets/97a346b3-178b-4a57-8b66-200efd2bd42e" />




<img width="1920" height="1080" alt="075" src="https://github.com/user-attachments/assets/8ded5094-bc7d-4ed1-b6c2-a105743e33e6" />




<img width="1920" height="1080" alt="053" src="https://github.com/user-attachments/assets/8fbf7c91-0cbb-4838-867a-84b7dcf41c2c" />






<img width="1920" height="1080" alt="038" src="https://github.com/user-attachments/assets/a1006daa-cb10-4ae6-a1c9-87d365025333" />





<img width="1920" height="1080" alt="037" src="https://github.com/user-attachments/assets/1735a0bc-9d00-435a-a1a3-17e31bc443b4" />



<img width="1920" height="1080" alt="036" src="https://github.com/user-attachments/assets/0b26a009-9c2e-4ac0-b811-61ead16b80da" />








<img width="1920" height="1080" alt="026" src="https://github.com/user-attachments/assets/81db1586-871f-4f6a-852e-c377d8daf0ac" />


<img width="1920" height="1080" alt="015" src="https://github.com/user-attachments/assets/b936b9f4-1a73-4f2b-ba4c-eded97f51f58" />




<img width="1920" height="1080" alt="014" src="https://github.com/user-attachments/assets/228b128e-3a8a-4b54-9296-16a97c96e5cc" />




[ Table of Contents](#table-of-contents)

---
---
---
---













# YearBook




<img width="1920" height="1080" alt="090" src="https://github.com/user-attachments/assets/056425b3-45e9-4c3d-a14c-d1104bf20399" />


<img width="1920" height="1080" alt="011" src="https://github.com/user-attachments/assets/0b1bff67-26a9-456d-9c42-0b96d4dd1bb8" />


[ Table of Contents](#table-of-contents)

---
---
---
---


