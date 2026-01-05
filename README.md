# 📁 File Explorer

**File Explorer** is a simple and user-friendly file management app designed to provide an intuitive interface for navigating, copying, moving, and deleting files and directories.




<img width="1266" height="662" alt="039" src="https://github.com/user-attachments/assets/7419a467-9180-4803-aa98-2f8f4f2de7f5" />

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


The **Command Line Interface (CLI)** for File Explorer application provides users with a way to interact with the file management system through text-based commands. This interface allows for efficient navigation and file operations, catering to users who prefer command-line interactions over graphical interfaces.

## Key Features of the CLI

###  Basic Commands
The CLI supports several basic commands that enable users to perform common file operations:

- **Change Directory (`cd`)** : 
  - Usage: `cd [directory]`
  - Changes the current working directory to the specified path.
  - Example: `cd C:\`



- **Make Directory (`mkdir` or `make`)** :

  - Usage: `mkdir [directory]`
  - Creates a new directory at the specified path.
  - Example: `mkdir C:\newfolder`





- **Copy Files (`copy`)** : 
  - Usage: `copy [source] [destination]`
  - Copies files or directories from a source to a destination.
  - Example: `copy C:\folder1\file.txt C:\folder2`

<img width="1266" height="662" alt="024" src="https://github.com/user-attachments/assets/ff0a34f8-ffe5-4fde-b2a5-d6615bb3faea" />



<img width="1266" height="662" alt="025" src="https://github.com/user-attachments/assets/e51e69a2-5ea4-4aea-8d1d-e01ab917b10b" />

- **Move Files (`move`)** : 
  - Usage: `move [source] [destination]`
  - Moves files or directories from one location to another.
  - Example: `move C:\folder1\file.txt C:\folder2\file.txt`

- **Delete Files (`delete`)** : 
  - Usage: `delete [file or directory]`
  - Deletes the specified file or directory.
  - Example: `delete C:\folder1\file.txt`
 
- **Create Text File (`text` or `txt`)** :
  - Usage: `text [file]`
  - Creates a new text file at the specified path and opens it.
  - Example: `text C:\folder\example.txt`
 


<img width="1266" height="662" alt="033" src="https://github.com/user-attachments/assets/5d4d9beb-8345-41db-a5dd-640e06427117" />


<img width="1266" height="662" alt="034" src="https://github.com/user-attachments/assets/669f727f-a3d5-468d-9457-77846120f02f" />


- **Help (`help`)** : 
  - Displays a list of available commands and their usage.
  - Example: `help`

###  Command Execution
- Users enter commands in a text box and press **Enter** to execute them.
- The application parses the command input, using regular expressions to handle quoted strings and spaces correctly.

###  Feedback and Status Messages
- The CLI provides immediate feedback on command execution, displaying messages for success, errors, or usage instructions.
- Example feedback messages include:
  - "No command entered."
  - "Copy Failed - Source does not exist."
  - "Deleted file: [file_path]"

###  Navigation History
- The CLI maintains a simple navigation history, allowing users to move back and forth through their directory changes.
- Users can navigate to previously accessed directories using the history feature, enhancing usability.

###  Contextual Commands
- If a command does not match any predefined operation, the CLI checks if the input corresponds to an existing file or directory and attempts to navigate to it or open it accordingly.

### Example Usage
Here’s how a typical session might look in the CLI:

```plaintext
> cd C:\Users\Joe
Navigated To: C:\Users\Joe

> copy C:\Users\Joe\file.txt C:\Users\Joe\Documents
Copied file: file.txt to: C:\Users\Joe\Documents

> delete C:\Users\Joe\Documents\old_file.txt
Deleted file: C:\Users\Joe\Documents\old_file.txt

> help
Available Commands:
cd [directory] - Change directory
copy [source] [destination] - Copy file or folder to destination folder
move [source] [destination] - Move file or folder to destination
delete [file_or_directory] - Delete file or directory
help - Show this help message
```


The CLI in the File Explorer application provides a powerful and flexible way to manage files and directories. It appeals to users who are comfortable with command-line operations, offering an efficient alternative to the graphical user interface. With support for essential file operations, feedback mechanisms, and navigation history, the CLI enhances the overall user experience.

---




---


# 📁 `RenameFileOrDirectory` —  Walkthrough

This method renames either a **file** or a **folder**, following a clear set of safety rules. Each rule protects the user from mistakes and helps them understand what’s happening.

<img width="1275" height="662" alt="040" src="https://github.com/user-attachments/assets/db463d52-0d3f-40c7-bb34-867deb1925d7" />

Below is the full code, then we’ll walk through it one small step at a time.

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



---

## 🔧 Method Signature

```vb.net
Private Sub RenameFileOrDirectory(sourcePath As String, newName As String)
```

- **sourcePath** — the full path to the file or folder you want to rename  
- **newName** — just the new name (not a full path)

---

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

---

## 🛑 Rule 1 — Path must be absolute

```vb.net
If Not Path.IsPathRooted(sourcePath) Then
    ShowStatus(IconDialog & " Rename failed: Path must be absolute. Example: C:\folder")
    Exit Sub
End If
```

Beginners often type relative paths like `folder\file.txt`.  
This rule stops the rename and explains the correct format.

---

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

---

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

---

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

---

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

---

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

---

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

---

This method teaches six important rules:

1. **Paths must be absolute**  
2. **Protected paths cannot be renamed**  
3. **Folders are renamed and then shown**  
4. **Files are renamed and their folder is shown**  
5. **Missing paths trigger a helpful quoting explanation**  
6. **Any error is caught and explained safely**









<img width="1275" height="662" alt="041" src="https://github.com/user-attachments/assets/15de34d4-f14c-4968-a10d-2203af0c7130" />





## License
This project is licensed under the **MIT License**. You are free to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the software, with the inclusion of the copyright notice and permission notice in all copies or substantial portions of the software.

## Installation
To install and run the File Explorer application:
1. Clone the repository:
   ```bash
   git clone https://github.com/JoeLumbley/File-Explorer.git
   ```
2. Open the solution in Visual Studio.
3. Build the project and run it.

## Usage
- Launch the application to access your file system.
- Use the tree view on the left to navigate through folders.
- The list view on the right displays the contents of the selected directory.
- Use the text box to enter commands or navigate directly to a path.
- Right-click on files or folders for additional options.


## Acknowledgements
This project is inspired by traditional file explorers and aims to provide a simplified experience for managing files on Windows systems.

For more details, check the source code and documentation within the repository.






<img width="1920" height="1080" alt="038" src="https://github.com/user-attachments/assets/a1006daa-cb10-4ae6-a1c9-87d365025333" />





<img width="1920" height="1080" alt="037" src="https://github.com/user-attachments/assets/1735a0bc-9d00-435a-a1a3-17e31bc443b4" />



<img width="1920" height="1080" alt="036" src="https://github.com/user-attachments/assets/0b26a009-9c2e-4ac0-b811-61ead16b80da" />








<img width="1920" height="1080" alt="026" src="https://github.com/user-attachments/assets/81db1586-871f-4f6a-852e-c377d8daf0ac" />


<img width="1920" height="1080" alt="015" src="https://github.com/user-attachments/assets/b936b9f4-1a73-4f2b-ba4c-eded97f51f58" />




<img width="1920" height="1080" alt="014" src="https://github.com/user-attachments/assets/228b128e-3a8a-4b54-9296-16a97c96e5cc" />








<img width="1920" height="1080" alt="011" src="https://github.com/user-attachments/assets/0b1bff67-26a9-456d-9c42-0b96d4dd1bb8" />




