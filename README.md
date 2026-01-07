# 📁 File Explorer

**File Explorer** is a simple and user-friendly file management app designed to provide an intuitive interface for navigating, copying, moving, and deleting files and directories.



<img width="1184" height="662" alt="052" src="https://github.com/user-attachments/assets/5eb1a283-6907-45a1-a605-6bd626519ac2" />


## Features
- **Navigation**: Easily navigate through directories using a tree view and a list view.
- **File Operations**: Perform common file operations such as:
  - Change Directory (`cd`)
  - Copy files and directories
  - Move files and directories
  - Delete files and directories
- **Clipboard Support**: Supports cut, copy, and paste operations for files and directories.
- **Context Menu**: Right-click context menu for quick access to file operations.
- **History Tracking**: Simple navigation history allows users to go back and forward through their directory changes.
- **File Type Icons**: Visual representation of different file types with corresponding icons.
- **Status Bar**: Provides feedback on the current operation and status updates.








---
---
---
---





## Table of Contents

- [🖥️ Command Line Interface (CLI)](#-command-line-interface)
  - [Features Overview](#-features-overview)
  - [Commands](#-commands)
    - [cd — Change Directory](#-change-directory--cd)
    - [mkdir / make — Create Directory](#-create-directory--mkdir-make)
    - [copy — Copy Files or Folders](#-copy--copy)
    - [move — Move Files or Folders](#-move--move)
    - [delete — Delete Files or Folders](#-delete--delete)
    - [rename — Rename Files or Folders](#-rename--rename)
    - [text / txt — Create Text Files](#-create-text-file--text-txt)
    - [find / search — Search](#-search--find-search)
    - [findnext / searchnext — Next Search Result](#-next-search-result--findnext-searchnext)
    - [help — Show Help](#-help--help)
  - [Quoting Rules](#-quoting-rules-important)
  - [Contextual Navigation](#-contextual-navigation)
  - [Example Session](#-example-session)

- [📦 MoveFileOrDirectory — Code Walkthrough](#movefileordirectory--code-walkthrough)

- [✏️ RenameFileOrDirectory — Code Walkthrough](#renamefileordirectory--code-walkthrough)

- [🧬 Clones](#clones)

- [📄 License](#license)






---
---
---
---










# 🖥️ Command Line Interface

The **Command Line Interface (CLI)** is an integrated text‑based command system inside the File Explorer application. It allows users to navigate folders, manage files, and perform common operations quickly using typed commands.

The CLI is designed to be:

- **Fast** — no menus, no dialogs  
- **Predictable** — clear rules and consistent behavior  
- **Beginner‑friendly** — helpful messages and examples  
- **Powerful** — supports navigation, search, file operations, and more  

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

---

## ⏭ Next Search Result — `findnext`, `searchnext`

**Usage:**  
```
findnext
```

**Description:**  
Cycles to the next result from the previous search.  
Wraps around when reaching the end.

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

---

# 🧭 Contextual Navigation

If the user enters something that is **not** a command:

- If it’s a **folder path**, the CLI navigates to it  
- If it’s a **file path**, the CLI opens it  
- Otherwise, the CLI shows an “Unknown command” message  

This makes the CLI feel natural and forgiving.

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




<img width="1266" height="662" alt="050" src="https://github.com/user-attachments/assets/60446ac8-b44a-447c-bcd5-4bffb25e2e8d" />


Below is the full code, then we’ll walk through it one small step at a time.


```vb.net

    Private Sub CopyDirectory(sourceDir As String, destDir As String)

        Dim dirInfo As New DirectoryInfo(sourceDir)

        If Not dirInfo.Exists Then

            ShowStatus(IconError & "  Source directory not found: " & sourceDir)

            Exit Sub

        End If

        Try

            ShowStatus(IconCopy & "  Create destination directory:" & destDir)

            ' Create destination directory
            Directory.CreateDirectory(destDir)

            ShowStatus(IconCopy & "  Copying files to destination directory:" & destDir)

            ' Copy files
            For Each file In dirInfo.GetFiles()
                Dim targetFilePath = Path.Combine(destDir, file.Name)
                file.CopyTo(targetFilePath, overwrite:=True)
            Next

            ShowStatus(IconCopy & "  Copying subdirectories.")

            ' Copy subdirectories recursively
            For Each subDir In dirInfo.GetDirectories()
                Dim newDest = Path.Combine(destDir, subDir.Name)
                CopyDirectory(subDir.FullName, newDest)
            Next

            ' Refresh the view to show the copied directory
            NavigateTo(destDir)

            ShowStatus(IconSuccess & "  Copied into " & destDir)

        Catch ex As Exception
            ShowStatus(IconError & "  Copy failed: " & ex.Message)
            Debug.WriteLine("CopyDirectory Error: " & ex.Message)
        End Try

    End Sub



```




## 🔧 Method Definition

```vb.net
Private Sub CopyDirectory(sourceDir As String, destDir As String)
```

- **sourceDir** — the folder you want to copy  
- **destDir** — where the copy should be created  



## Create a DirectoryInfo object for the source

```vb.net
Dim dirInfo As New DirectoryInfo(sourceDir)
```

- `DirectoryInfo` gives you access to:
  - the folder’s files  
  - its subfolders  
  - metadata  
- It’s a convenient wrapper around a directory path.



## Make sure the source directory exists

```vb.net
If Not dirInfo.Exists Then
    ShowStatus(IconError & "  Source directory not found: " & sourceDir)
    Exit Sub
End If
```

- If the folder doesn’t exist, we stop immediately.  
- Beginners often mistype paths, so this prevents confusing errors.  
- The user gets a clear, friendly message.



## Start a Try/Catch block

```vb.net
Try
```

Everything inside this block is protected.  
If anything goes wrong (permissions, locked files, etc.), the `Catch` block will handle it gracefully.



## Tell the user we’re creating the destination directory

```vb.net
ShowStatus(IconCopy & "  Create destination directory:" & destDir)
```

This gives immediate feedback so the UI feels alive and responsive.



## Create the destination directory

```vb.net
Directory.CreateDirectory(destDir)
```

- If the folder already exists, nothing bad happens.  
- If it doesn’t exist, it is created.  
- Either way, the destination is now ready.



## Tell the user we’re copying files

```vb.net
ShowStatus(IconCopy & "  Copying files to destination directory:" & destDir)
```

This message helps user understand the sequence of operations.



## Copy all files in the current directory

```vb.net
For Each file In dirInfo.GetFiles()
    Dim targetFilePath = Path.Combine(destDir, file.Name)
    file.CopyTo(targetFilePath, overwrite:=True)
Next
```

- `GetFiles()` returns all files directly inside the folder.  
- `Path.Combine` builds the full destination path.  
- `CopyTo(..., overwrite:=True)` ensures:
  - files are copied  
  - existing files are replaced  

This loop handles only the files — not subfolders.



## Tell the user we’re copying subdirectories

```vb.net
ShowStatus(IconCopy & "  Copying subdirectories.")
```

This prepares the user for the next step: recursion.



## Copy all subdirectories (recursively)

```vb.net
For Each subDir In dirInfo.GetDirectories()
    Dim newDest = Path.Combine(destDir, subDir.Name)
    CopyDirectory(subDir.FullName, newDest)
Next
```

This is the heart of the algorithm.

- `GetDirectories()` returns all subfolders.  
- For each subfolder:
  - Build a new destination path  
  - Call **CopyDirectory again** on that subfolder  

This technique is called **recursion** — the function keeps calling itself until it reaches the deepest level of the folder tree.

Beginners often find this magical once they see it in action.



## Refresh the UI to show the copied directory

```vb.net
NavigateTo(destDir)
```

This helps the user visually confirm the copy succeeded.



## Show a success message

```vb.net
ShowStatus(IconSuccess & "  Copied into " & destDir)
```

Clear, friendly confirmation that the operation completed.



## Handle any errors

```vb.net
Catch ex As Exception
    ShowStatus(IconError & "  Copy failed: " & ex.Message)
    Debug.WriteLine("CopyDirectory Error: " & ex.Message)
End Try
```

If anything goes wrong:

- The user gets a helpful message  
- You get a debug log for troubleshooting  

This keeps the app stable and user‑friendly.



This method shows:

- How to check whether a directory exists  
- How to create directories safely  
- How to copy files  
- How to copy subfolders using **recursion**  
- How to build paths correctly  
- How to give user feedback  
- How to handle errors without crashing  



<img width="1266" height="662" alt="051" src="https://github.com/user-attachments/assets/1ec25af2-62d5-4877-a9d6-4210342ae4e3" />



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


<img width="1920" height="1080" alt="075" src="https://github.com/user-attachments/assets/8ded5094-bc7d-4ed1-b6c2-a105743e33e6" />




<img width="1920" height="1080" alt="053" src="https://github.com/user-attachments/assets/8fbf7c91-0cbb-4838-867a-84b7dcf41c2c" />






<img width="1920" height="1080" alt="038" src="https://github.com/user-attachments/assets/a1006daa-cb10-4ae6-a1c9-87d365025333" />





<img width="1920" height="1080" alt="037" src="https://github.com/user-attachments/assets/1735a0bc-9d00-435a-a1a3-17e31bc443b4" />



<img width="1920" height="1080" alt="036" src="https://github.com/user-attachments/assets/0b26a009-9c2e-4ac0-b811-61ead16b80da" />








<img width="1920" height="1080" alt="026" src="https://github.com/user-attachments/assets/81db1586-871f-4f6a-852e-c377d8daf0ac" />


<img width="1920" height="1080" alt="015" src="https://github.com/user-attachments/assets/b936b9f4-1a73-4f2b-ba4c-eded97f51f58" />




<img width="1920" height="1080" alt="014" src="https://github.com/user-attachments/assets/228b128e-3a8a-4b54-9296-16a97c96e5cc" />








<img width="1920" height="1080" alt="011" src="https://github.com/user-attachments/assets/0b1bff67-26a9-456d-9c42-0b96d4dd1bb8" />




