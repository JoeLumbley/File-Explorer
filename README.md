# File Explorer


**File Explorer** is a simple, fast, and user‑friendly file management application designed to make navigating, organizing, and manipulating files intuitive for all users. It combines a clean graphical interface with a powerful built‑in **Command Line Interface (CLI)** for those who prefer keyboard‑driven workflows.


<img width="1266" height="733" alt="118" src="https://github.com/user-attachments/assets/eddcf5af-98f4-406c-a742-439325ccee10" />

The application is built around clarity, predictability, and emotional safety. Whether you prefer clicking, typing, or shortcut‑driven workflows, File Explorer adapts to your style.

[ Table of Contents](#table-of-contents)


## **Features**

### **Graphical Interface**  
• Browse directories using a tree view and list view.  
• Perform file operations such as copy, move, delete, and rename.  
• Use cut, copy, and paste for file and folder management.  
• Access context menus for quick actions.  
• Navigate backward and forward through history.  
• View file type icons and real‑time status updates.  
• Open the **Help Drawer**, a panel with beginner‑friendly documentation.

### **Integrated Command Line Interface (CLI)**  
• Fast directory navigation (`cd`).  
• File operations (`copy`, `move`, `delete`, `rename`).  
• Create text files (`text`, `txt`).  
• Search and cycle results (`find`, `findnext`).  
• Type a path to open it directly.  
• Supports quoted paths with spaces.  
• Helpful usage messages and error feedback.  
• Built‑in help system with full documentation in the Help Drawer.

<img width="1266" height="733" alt="114" src="https://github.com/user-attachments/assets/a00cf73a-91b0-43be-a94f-1cee53d79fd3" />


The GUI and CLI work together seamlessly, giving users the freedom to choose the workflow that suits them best.



[ Table of Contents](#table-of-contents)


---
---
---
---







## Why I’m Creating File Explorer

I set out to build my own File Explorer because I wanted to understand, from the ground up, how a core part of every operating system actually works. We all use file managers every day, but it’s easy to overlook how much is happening behind the scenes: navigation history, sorting, file type detection, context menus, clipboard operations, lazy loading of folder trees, and many other details. Recreating these features myself has been a practical way to explore system I/O, UI design, event handling, and performance considerations in a hands-on, exploratory way.

This project is not meant to replace the built-in Windows Explorer. Instead, it serves as a learning environment, a place where I can experiment, break things, fix them, and understand why they work the way they do. By rebuilding something familiar, I get to uncover the subtle engineering decisions that make everyday tools feel intuitive, similar to how art students copy the masters to study technique and intention.


<img width="1920" height="1080" alt="105" src="https://github.com/user-attachments/assets/2144557d-8c45-4278-b2fb-f485ab5f2212" />





[ Table of Contents](#table-of-contents)




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






### **A space for deliberate practice and student‑driven growth**
One of the core ideas behind this project is the value of deliberate practice, breaking a complex system into understandable pieces, studying them closely, and rebuilding them with intention. File Explorer is a perfect playground for that kind of learning. It’s small enough to grasp, but rich enough to teach real engineering habits: decomposition, naming, event flow, UI state management, and the discipline of making things predictable for users.
My hope is that this project becomes a starting point for students who want to build their own great applications. By exploring the code, modifying features, or adding entirely new ones, learners can practice the same skills professional developers use every day. This isn’t just a tool to look at, it’s a foundation you can extend, reshape, and eventually outgrow as you build projects of your own.



If you’re curious, the GitHub repository includes the full source code and documentation. I’d love to hear your thoughts, suggestions, or ideas for future features. This project is as much about learning as it is about building something functional, and I’m excited to share that journey with you.



[ Table of Contents](#table-of-contents)





---
---
---
---





















# Command Line Interface

The **Command Line Interface (CLI)** is an integrated text‑based command system inside the File Explorer application. It allows users to navigate folders, manage files, and perform common operations quickly using typed commands.




<img width="1266" height="733" alt="119" src="https://github.com/user-attachments/assets/bad79464-9688-4502-b511-ba8cef03b0fd" />





The CLI is designed to be:

- **Fast** - no menus, no dialogs  
- **Predictable** - clear rules and consistent behavior  
- **Beginner‑friendly** - helpful messages and examples  
- **Powerful** - supports navigation, search, file operations, and more  

[ Table of Contents](#table-of-contents)

---

##  Features Overview

###  Navigation  
- Change directories  
- Open files directly  
- Navigate to folders by typing their path  
- Supports paths with spaces using quotes  

###  File & Folder Operations  
- Create, copy, move, rename, and delete  
- Works with both files and directories  
- Handles quoted paths safely  

###  Search  
- Search the current folder  
- Cycle through results with `findnext`  
- Highlights and selects results in the UI  

###  Contextual Behavior  
If a command doesn’t match a known keyword, the CLI checks:

- **Is it a folder?** → Navigate to it  
- **Is it a file?** → Open it  
- Otherwise → “Unknown command”  

This makes the CLI feel natural and forgiving.

[ Table of Contents](#table-of-contents)

---




# **Commands**

## **Command Index**

- [`cd`](#-change-directory--cd)  
- [`copy cp`](#-copy--copy-cp)  
- [`delete rm`](#-delete--delete-rm)  
- [`df`](#-disk-free--df)  
- [`drives`](#-drives-overview--drives)  
- [`exit quit close`](#-exit--exit-quit-close-stop-halt-end-signout-poweroff-bye)  
- [`find search`](#-search--find-search)  
- [`findnext searchnext next`](#-next-search-result--findnext-searchnext-next)  
- [`help commands ?`](#-help--help-commands-)  
- [`man manual appmanual`](#-manual--man-manual-appmanual)  
- [`mkdir make md`](#-create-directory--mkdir-make-md)  
- [`move mv`](#-move--move-mv)  
- [`open`](#-open--open)  
- [`pin`](#-pin--pin)  
- [`rename rn`](#-rename--rename-rn)  
- [`shortcuts keys`](#-keyboard-shortcuts--shortcuts-keys)  
- [`text txt`](#-create-text-file--text-txt)  

[ Table of Contents](#table-of-contents)

Below is the complete list of supported commands, including syntax, descriptions, and examples.

---

## 📁 Create Directory — `mkdir`, `make`, `md`

**Usage**
```
mkdir [directory_path]
```

**Description**  
Create a new folder at the specified path.

**Examples**
```
mkdir C:\newfolder
make "C:\My New Folder"
md C:\anotherfolder
```
[Command Index](#command-index)

---

## 📌 Pin — `pin`

**Usage**
```
pin [folder_path]
```

**Description**  
Pin or unpin a folder.  
If no path is provided, the current folder is used when valid.

**Examples**
```
pin C:\Projects
pin "C:\My Documents"
pin
```
[Command Index](#command-index)

---

## 📄⇢📁 Copy — `copy`, `cp`

**Usage**
```
copy [source] [destination]
```

**Description**  
Copy a file or folder to the specified destination.

**Examples**
```
copy C:\folderA\file.doc C:\folderB
copy "C:\folder A" "C:\folder B"
```
[Command Index](#command-index)

---

## 📦 Move — `move`, `mv`

**Usage**
```
move [source] [destination]
```

**Description**  
Move a file or folder to a new location.

**Examples**
```
move C:\folderA\file.doc C:\folderB\file.doc
move "C:\folder A\file.doc" "C:\folder B\renamed.doc"
```
[Command Index](#command-index)

---

## ✏ Rename — `rename`, `rn`

**Usage**
```
rename [source_path] [new_name]
```

**Description**  
Rename a file or folder.

**Examples**
```
rename "C:\folder\oldname.txt" "newname.txt"
```
[Command Index](#command-index)

---

## 🗑 Delete — `delete`, `rm`

**Usage**
```
delete [file_or_directory]
```

**Description**  
Delete the specified file or folder.

**Examples**
```
delete C:\file.txt
delete "C:\My Folder"
```
[Command Index](#command-index)

---

## 🔍 Search — `find`, `search`

**Usage**
```
find [search_term]
```

**Description**  
Search the current directory for items whose names contain the given term.

**Examples**
```
find document
```
[Command Index](#command-index)

---

## ⏭ Next Search Result — `findnext`, `searchnext`, `next`

**Usage**
```
findnext
```

**Description**  
Show the next result from the previous search.

---

## 📁 Change Directory — `cd`

**Usage**
```
cd [directory]
```

**Description**  
Change the current working directory.

**Examples**
```
cd C:\
cd "C:\My Folder"
```
[Command Index](#command-index)

---

## 📂 Open — `open`

**Usage**
```
open [file_or_directory]
```

**Description**  
Open a file or navigate into a folder.

**Examples**
```
open C:\folder\file.txt
open "C:\My Folder"
```
[Command Index](#command-index)

---

## 📝 Create Text File — `text`, `txt`

**Usage**
```
text [file_path]
```

**Description**  
Create a new text file at the specified path.

**Examples**
```
text "C:\folder\example.txt"
```
[Command Index](#command-index)

---

## 💽 Disk Free — `df`

**Usage**
```
df <drive_letter>:
```

**Description**  
Display free and total space for the specified drive.

**Examples**
```
df C:
df D:
df E:
```
[Command Index](#command-index)

---

## 💾 Drives Overview — `drives`

**Usage**
```
drives
```

**Description**  
Show all available drives along with their free‑space information.

**Examples**
```
drives
```
[Command Index](#command-index)

---

## ❓ Help — `help`, `commands`, `?`

**Usage**
```
help [search_term]
```

**Description**  
Show the full list of commands or display help for a specific command.

**Examples**
```
help
help cd
help copy
```
[Command Index](#command-index)

---

## 📘 Manual — `man`, `manual`, `appmanual`

**Usage**
```
man [section]
```

**Description**  
Show the full application manual or jump to a specific section.

**Examples**
```
man
man help
man commands
manual
appmanual
```
[Command Index](#command-index)

---

## ⌨ Keyboard Shortcuts — `shortcuts`, `keys`

**Usage**
```
shortcuts
```

**Description**  
Show all available keyboard shortcuts.

**Examples**
```
shortcuts
keys
```
[Command Index](#command-index)

---

## ❌ Exit — `exit`, `quit`, `close`, `stop`, `halt`, `end`, `signout`, `poweroff`, `bye`

**Usage**
```
exit
```

**Description**  
Close the application.

---


[Command Index](#command-index)








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

It’s a flexible, efficient alternative to the graphical interface - perfect for users who enjoy command‑driven workflows.


[ Table of Contents](#table-of-contents)








---
---
---
---
























---
---
---
---
















## Table of Contents

- [Why I’m Creating File Explorer](#why-im-creating-file-explorer)
  
- [What I Hope Learners Get From This](#what-i-hope-learners-get-from-this)

- [Code Walkthrough](#-code-walkthrough-systems-that-make-up-the-app)


- [Command Line Interface (CLI)](#command-line-interface)
  - [Features Overview](#features-overview)
  - [Commands](#commands)
  - [Quoting Rules](#-quoting-rules-important)
  - [Contextual Navigation](#-contextual-navigation)
  - [Example Session](#-example-session)

- [⌨️ Keyboard Shortcuts](#keyboard-shortcuts)


- [📄 License](#license)

- [🧬 Clones](#clones)

- [📚 YearBook](#yearbook)


[Back to the start](#file-explorer)

---
---
---
---










































































# 🧩 Code Walkthrough: Systems That Make Up the App

This application is built as a set of small, focused systems that work together.  
In this walkthrough, we go **line by line** through the core systems that make up the app and see how each one is implemented in code.

---

## **Systems Index**

- [Directory & File Navigation System](#directory--file-navigation-system)  
- [Command Line Interface (CLI) System](#command-line-interface-cli-system)  
- [Help & Manual System](#help--manual-system)  
- [Keyboard Input & Routing System](#keyboard-input--routing-system)  
- [File Operation System](#file-operation-system)  
- [Search System](#search-system)  
- [GUI Rendering System](#gui-rendering-system)  
- [Pinning System](#pinning-system)  
- [State Management System](#state-management-system)  
- [Exit & Safety System](#exit--safety-system)  

[ Table of Contents](#table-of-contents)

[Back to the start](#file-explorer)

---

### Directory & File Navigation System

This system is responsible for **where you are** and **what you see**.

- Tracks the current directory and updates the UI when it changes  
- Keeps the **tree view** and **list view** in sync  
- Manages history (Back/Forward) and selection state  
- Resolves paths (absolute, relative, quoted)

In the code walkthrough, we look at:

- How the current path is stored  
- How directory changes trigger UI refreshes  
- How selection is read and used by commands like `open`, `delete`, and `rename`

[Systems Index](#systems-index)

---

### Command Line Interface (CLI) System

This system turns user input into actions.

- Parses commands and arguments  
- Resolves aliases using the `CommandHelp` dictionary  
- Handles quoted paths and spacing rules  
- Routes each command to the correct handler method  
- Shows usage and error messages when something isn’t valid

In the walkthrough, we go through:

- The main input handler  
- How the command name is matched against `CommandHelp`  
- How arguments are split, validated, and passed to operations

[Systems Index](#systems-index)

---

### Help & Manual System

This system makes the app **self‑documenting**.

- `CommandHelp` dictionary with aliases, usage, descriptions, and examples  
- `help` command for command‑level help  
- `man` command for full manual sections  
- Rendering logic for help text and manual pages  
- Search and fallback behavior when an exact match isn’t found

In the walkthrough, we examine:

- How `CommandHelp` is built  
- How `BuildHelpText` constructs the help output  
- How `help` and `man` decide what to show

[Systems Index](#systems-index)

---

### Keyboard Input & Routing System

This system decides **what a key press means** in each context.

- Handles Enter, Escape, Tab, Shift+Tab, and shortcuts  
- Routes keys differently depending on focus (CLI, list view, Help Drawer)  
- Suppresses unsafe repeats and protects modal states  
- Keeps behavior predictable and consistent

In the walkthrough, we look at:

- The central key handler  
- How context is detected (which control is active)  
- How routing tables or `Select Case` blocks map keys to actions

[Systems Index](#systems-index)

---

### File Operation System

This system performs the actual work on the filesystem.

- Create: `mkdir`, `text`  
- Copy: `copy`  
- Move: `move`  
- Delete: `delete`  
- Rename: `rename`  
- Open: `open`  
- Pin/Unpin: `pin`  
- Disk info: `df`, `drives`

In the walkthrough, we step through:

- How each command validates paths  
- How errors are handled and surfaced to the user  
- How the UI is refreshed after an operation completes

[Systems Index](#systems-index)

---

### Search System

This system finds items in the current directory.

- `find` to start a search  
- `findnext` to move through results  
- Stores the current result list and index  
- Automatically selects the first match  
- Wraps around when reaching the end

In the walkthrough, we cover:

- How the search term is applied to items  
- How results are stored and reused  
- How selection is updated as you cycle through matches

[Systems Index](#systems-index)

---

### GUI Rendering System

This system draws what the user sees.

- Tree view and list view population  
- Status bar updates (current path, selection, search results)  
- Icons and visual indicators  
- Help Drawer layout and content  
- Context menus and their actions

In the walkthrough, we look at:

- How the UI is initialized  
- How data binding or manual population works  
- How the GUI reacts to changes from the CLI

[Systems Index](#systems-index)

---

### Pinning System

The pinning system manages the set of folders the user has marked for quick access. It provides a simple toggle‑based workflow and keeps both the internal state and the user interface consistent.

This system is responsible for:

- Tracking which folders are currently pinned  
- Implementing the toggle behavior for the `pin` command  
- Validating paths before pinning or unpinning  
- Updating any UI elements that display pinned folders  
- Keeping pinned state synchronized across CLI, GUI, and internal state  

In the walkthrough, we examine:

- How pinned folders are stored, whether in memory or persisted to disk  
- How the `pin` command interacts with the pin list  
- How the toggle logic determines whether a folder should be pinned or unpinned  
- How invalid or special folders are rejected safely  
- How the UI is refreshed after pinning changes  
- How the system integrates with contextual selection  

[Pinning System – Walkthrough](#pinning-system---code-walkthrough)


[Systems Index](#systems-index)

---

### State Management System

This system keeps everything in sync.

- Current directory  
- Current selection  
- Search state  
- Help/manual state  
- Pinned folders  
- History

In the walkthrough, we examine:

- Where shared state lives  
- How different systems read and update it  
- How we avoid inconsistent or unsafe states

[Systems Index](#systems-index)

---

### Exit & Safety System

This system handles closing the app safely.

- `exit` command and its aliases  
- Confirmation prompts  
- Cleanup logic before shutdown

In the walkthrough, we cover:

- How exit is triggered from CLI and GUI  
- How the confirmation dialog is shown  
- How the app ensures a clean shutdown


[Systems Index](#systems-index)





---
---
---
---
























































# **Pinning System - Code Walkthrough**

The pinning system manages the user’s Easy Access list. It stores pinned folders, validates them, updates the UI, and provides the toggle behavior used by both the GUI and CLI. This walkthrough explains each method in the system line by line so readers can understand how the feature works internally.

---

## **Pinning System Index**

- [RefreshPinUI](#refreshpinui)  
- [EnsureEasyAccessFile](#ensureeasyaccessfile)  
- [LoadEasyAccessEntries](#loadeasyaccessentries)  
- [AddToEasyAccess](#addtoeasyaccess)  
- [RemoveFromEasyAccess](#removefromeasyaccess)  
- [IsPinned](#ispinned)  
- [UpdatePinButtonState](#updatepinbuttonstate)  
- [UpdateFileListPinState](#updatefilelistpinstate)  
- [IsTreeNodePinnable](#istreenodepinnable)  
- [UpdateTreeContextMenu](#updatetreecontextmenu)  
- [IsSpecialFolder](#isspecialfolder)  
- [PinFromFiles / UnpinFromFiles](#pinfromfiles--unpinfromfiles)  
- [Pin_Click / Unpin_Click](#pin_click--unpin_click)  
- [TogglePin](#togglepin)  
- [GetPinnableTarget](#getpinnabletarget)  

[Systems Index](#systems-index)

[ Table of Contents](#table-of-contents)

---

















## **RefreshPinUI**

---

### **What this method does**

`RefreshPinUI` updates every UI element that depends on the current pin state. It ensures that the tree view, file list, and toolbar button all reflect the latest pinned/unpinned status after any change.

---

```vb.net

Private Sub RefreshPinUI()
    UpdateTreeRoots()
    UpdateFileListPinState()
    UpdatePinButtonState()
End Sub

```

---


### **How the method works**

- Calls `UpdateTreeRoots()` to rebuild the folder tree, including pinned roots at the top.  
- Calls `UpdateFileListPinState()` to update the file list’s context menu so the correct Pin/Unpin option is shown.  
- Calls `UpdatePinButtonState()` to update the toolbar pin/unpin button based on the current folder or selection.

Together, these updates ensure the interface stays synchronized with the underlying Easy Access data.

---

### **Why this method matters**

`RefreshPinUI` is the central refresh point for the entire pinning system. It ensures:

- The UI always reflects the correct pin state.  
- All views (tree, list, toolbar) stay consistent with each other.  
- Any pin/unpin action—whether from CLI, context menu, or toolbar—immediately updates the interface.  
- The user never sees stale or outdated pin indicators.

Without this method, the UI could easily fall out of sync with the underlying data.

---

[Pinning System Index](#pinning-system-index)

---






















## **EnsureEasyAccessFile**

### **EnsureEasyAccessFile Index**
- [What this method does](#what-this-method-does-5)  
- [How the method works](#how-the-method-works-5)  
- [Why this method matters](#why-this-method-matters-5)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`EnsureEasyAccessFile` guarantees that the Easy Access storage file exists before any pinning operation is performed. It creates both the directory and the file if they are missing, ensuring the rest of the system can safely read and write entries.

---

```vb.net


Private Sub EnsureEasyAccessFile()
    Dim dir = Path.GetDirectoryName(EasyAccessFile)
    If Not Directory.Exists(dir) Then Directory.CreateDirectory(dir)
    If Not IO.File.Exists(EasyAccessFile) Then IO.File.WriteAllText(EasyAccessFile, "")
End Sub

```

---


### **How the method works**

- Extracts the directory portion of the Easy Access file path.  
- Checks whether that directory exists; if not, it creates it.  
- Checks whether the Easy Access file exists; if not, it creates an empty file.  

This ensures the storage location is always valid and ready for use.

---

### **Why this method matters**

`EnsureEasyAccessFile` prevents:

- File‑not‑found exceptions  
- Directory‑not‑found exceptions  
- Corrupted or missing Easy Access storage  
- Inconsistent behavior between sessions  

It acts as the foundation for all pinning operations, ensuring the system always has a safe, predictable place to store pinned folder entries.

---

[Pinning System Index](#pinning-system-index)

---















## **LoadEasyAccessEntries**

### **LoadEasyAccessEntries Index**
- [What this method does](#what-this-method-does-6)  
- [How the method works](#how-the-method-works-6)  
- [Why this method matters](#why-this-method-matters-6)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`LoadEasyAccessEntries` reads the Easy Access storage file and converts each valid line into a `(Name, Path)` tuple. It ensures the file exists, parses each entry safely, and returns the authoritative list of pinned folders used throughout the app.

---

```vb.net

Private Function LoadEasyAccessEntries() As List(Of (Name As String, Path As String))
    EnsureEasyAccessFile()

    Dim list As New List(Of (String, String))

    For Each line In IO.File.ReadAllLines(EasyAccessFile)
        Dim entry = ParseEntry(line)
        If entry.HasValue Then
            ' Keep even if missing
            list.Add((entry.Value.Name, entry.Value.Path))
        End If
    Next

    Return list
End Function

```

---


### **How the method works**

- Calls `EnsureEasyAccessFile()` to guarantee the storage file and directory exist.  
- Creates a new list to hold the parsed `(Name, Path)` entries.  
- Reads every line from the Easy Access file.  
- Uses `ParseEntry` to convert each line into a structured tuple.  
- Adds the entry to the list **even if the folder no longer exists**.  
- Returns the completed list to the caller.

This method centralizes all loading logic so the rest of the system can rely on a clean, structured list of pinned folders.

---

### **Why this method matters**

`LoadEasyAccessEntries` ensures:

- The pinning system always works with a consistent, normalized list of entries.  
- Corrupted or malformed lines do not break the system.  
- Missing folders remain visible, preserving user intent and matching Explorer semantics.  
- All UI components (tree, list, toolbar) can rebuild their pinned state from a single source of truth.

It is the backbone of the pinning system’s data layer.

---

[Pinning System Index](#pinning-system-index)

---




















## **AddToEasyAccess**

### **AddToEasyAccess Index**
- [What this method does](#what-this-method-does-1)  
- [How the method works](#how-the-method-works-1)  
- [Why this method matters](#why-this-method-matters-1)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`AddToEasyAccess` adds a new folder to the Easy Access list. It ensures the storage file exists, prevents duplicates, writes the new entry, and refreshes the UI so the change is immediately visible.

---

```vb.net
Public Sub AddToEasyAccess(name As String, path As String)
    EnsureEasyAccessFile()

    Dim normalized = NormalizePath(path)
    Dim existing = IO.File.ReadAllLines(EasyAccessFile)

    ' Prevent duplicates by normalized path
    If existing.Any(Function(line)
                        Dim e = ParseEntry(line)
                        Return e.HasValue AndAlso NormalizePath(e.Value.Path) = normalized
                    End Function) Then
        RefreshPinUI()
        Exit Sub
    End If

    ' Write new entry
    IO.File.AppendAllLines(EasyAccessFile, {$"{name},{path}"})

    RefreshPinUI()
End Sub
```

---

### **How the method works**

- Ensures the Easy Access file exists before doing anything else.  
- Normalizes the incoming path so comparisons are consistent.  
- Reads all existing entries from the file.  
- Checks for duplicates by comparing normalized paths.  
- If the folder is already pinned, it simply refreshes the UI and exits.  
- If not pinned, it appends a new entry in `name,path` format.  
- Calls `RefreshPinUI()` to update the tree, file list, and pin button.

This prevents duplicate entries and keeps the UI synchronized with the underlying data.

---

### **Why this method matters**

`AddToEasyAccess` ensures:

- The pinned list stays clean and free of duplicates.  
- The Easy Access file is always valid and ready to use.  
- The UI updates immediately after any change.  
- The pinning system behaves consistently across CLI and GUI.

---

[Pinning System Index](#pinning-system-index)

---




















## **RemoveFromEasyAccess**

### **RemoveFromEasyAccess Index**
- [What this method does](#what-this-method-does-3)  
- [How the method works](#how-the-method-works-3)  
- [Why this method matters](#why-this-method-matters-3)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`RemoveFromEasyAccess` removes a folder from the Easy Access list. It ensures the storage file exists, filters out the target entry, writes the updated list back to disk, and refreshes the UI so the change is immediately visible.

---

```vb.net
Public Sub RemoveFromEasyAccess(path As String)
    EnsureEasyAccessFile()

    Dim target = NormalizePath(path)
    Dim lines = IO.File.ReadAllLines(EasyAccessFile)

    Dim updated = lines.Where(Function(line)
                                  Dim e = ParseEntry(line)
                                  If Not e.HasValue Then Return True
                                  Return NormalizePath(e.Value.Path) <> target
                              End Function).ToList()

    IO.File.WriteAllLines(EasyAccessFile, updated)

    RefreshPinUI()
End Sub

```

---

### **How the method works**

- Ensures the Easy Access file exists before doing anything else.  
- Normalizes the target path so comparisons are consistent.  
- Reads all lines from the Easy Access file.  
- Parses each entry and filters out any whose normalized path matches the target.  
- Writes the updated list back to the file.  
- Calls `RefreshPinUI()` to update the tree, file list, and pin button.

This cleanly removes the folder from the pinned list and ensures the UI reflects the change immediately.

---

### **Why this method matters**

`RemoveFromEasyAccess` ensures:

- Pinned folders can be safely and reliably removed.  
- The Easy Access file never contains stale or duplicate entries.  
- The UI stays synchronized with the underlying data.  
- The toggle behavior (`pin` / `unpin`) remains predictable and consistent.

---

[Pinning System Index](#pinning-system-index)

---
























---

## **IsPinned**

### **IsPinned Index**
- [What this method does](#what-this-method-does-7)  
- [How the method works](#how-the-method-works-7)  
- [Why this method matters](#why-this-method-matters-7)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`IsPinned` determines whether a given folder is currently pinned by checking the Easy Access storage file. It normalizes the path, parses each entry, and returns `True` if any entry matches the target folder.

---

```vb.net
Private Function IsPinned(path As String) As Boolean
    Dim target = NormalizePath(path)
    Return File.ReadAllLines(EasyAccessFile).
    Any(Function(line)
            Dim e = ParseEntry(line)
            Return e.HasValue AndAlso NormalizePath(e.Value.Path) = target
        End Function)
End Function

```

---

### **How the method works**

- Normalizes the input path to ensure consistent comparison.  
- Reads all lines from the Easy Access file.  
- Parses each line using `ParseEntry`.  
- Compares the normalized stored path against the normalized input path.  
- Returns `True` if a match is found; otherwise returns `False`.

This method is used throughout the pinning system to determine whether a folder should be pinned or unpinned, and to update UI elements accordingly.

---

### **Why this method matters**

`IsPinned` ensures:

- The toggle logic (`TogglePin`) behaves correctly.  
- The UI (tree view, list view, toolbar button) always reflects the correct pin state.  
- Duplicate entries are prevented by checking normalized paths.  
- The system remains consistent even if paths differ in casing or formatting.

It is a foundational helper method used across the entire pinning workflow.

---

[Pinning System Index](#pinning-system-index)

---

















## **UpdatePinButtonState**

### **UpdatePinButtonState Index**
- **What this method does**  
- **How the method works**  
- **Why this method matters**  
- **Back to Pinning System Index**

---

### **What this method does**

`UpdatePinButtonState` updates the toolbar’s pin/unpin button so it always reflects the correct state of the folder the user is currently interacting with. It determines whether a valid target exists, enables or disables the button accordingly, and switches the icon based on whether that folder is pinned.

---

```vb.net
Private Sub UpdatePinButtonState()
    btnPin.Enabled = False
    btnPin.Text = ""   ' Pin icon

    Dim target As String = GetPinnableTarget()
    If target Is Nothing Then Exit Sub

    btnPin.Enabled = True
    btnPin.Text = If(IsPinned(target), "", "")
End Sub


```

---

### **How the method works**

- It begins by disabling the button and setting the default “pin” icon. This prevents misleading UI states when no valid folder is selected.  
- It calls `GetPinnableTarget()` to determine which folder—if any—is currently eligible for pinning or unpinning.  
- If no valid target is returned, the method exits, leaving the button disabled.  
- If a valid folder is found, the button is enabled.  
- The icon is then updated:  
  - The “unpin” icon is shown if the folder is already pinned.  
  - The “pin” icon is shown if the folder is not pinned.

This ensures the toolbar button always communicates the correct action to the user.

---

### **Why this method matters**

`UpdatePinButtonState` keeps the toolbar intuitive and trustworthy. It ensures:

- The user never sees an enabled pin button when no folder can be pinned.  
- The icon always matches the actual state of the selected or active folder.  
- The UI stays synchronized with the underlying pin data and the user’s current context.  
- Keyboard, mouse, and CLI interactions all produce consistent visual feedback.

Without this method, the toolbar could easily fall out of sync with the rest of the interface.

---

[Pinning System Index](#pinning-system-index)

---
























## **UpdateFileListPinState**

### **UpdateFileListPinState Index**
- **What this method does**  
- **How the method works**  
- **Why this method matters**  
- **Back to Pinning System Index**

---

### **What this method does**

`UpdateFileListPinState` updates the visibility of the **Pin** and **Unpin** options in the file list’s context menu. It evaluates the currently selected item, determines whether it is a valid folder, and shows the appropriate menu option based on whether that folder is already pinned.

---

```vb.net
Private Sub UpdateFileListPinState()
    Dim mnuPin = cmsFiles.Items("Pin")
    Dim mnuUnpin = cmsFiles.Items("Unpin")

    mnuPin.Visible = False
    mnuUnpin.Visible = False

    If lvFiles.SelectedItems.Count = 0 Then Exit Sub

    Dim path As String = TryCast(lvFiles.SelectedItems(0).Tag, String)
    If String.IsNullOrEmpty(path) Then Exit Sub
    If Not Directory.Exists(path) Then Exit Sub
    If IsSpecialFolder(path) Then Exit Sub

    mnuPin.Visible = Not IsPinned(path)
    mnuUnpin.Visible = IsPinned(path)
End Sub

```

---

### **How the method works**

- Retrieves the `Pin` and `Unpin` menu items from the file list context menu.  
- Hides both options by default to avoid showing incorrect actions.  
- Exits early if no item is selected in the file list.  
- Extracts the selected item’s path and validates that:  
  - The path is not empty.  
  - The directory exists.  
  - The folder is not a special folder (Documents, Desktop, etc.).  
- If the folder is valid, it checks whether it is pinned:  
  - Shows **Pin** when the folder is not pinned.  
  - Shows **Unpin** when the folder is pinned.

This ensures the context menu always presents the correct action for the selected folder.

---

### **Why this method matters**

`UpdateFileListPinState` ensures:

- The user never sees both Pin and Unpin at the same time.  
- Invalid or non-folder items never show pin options.  
- The context menu stays synchronized with the actual pinned state.  
- The file list behaves consistently with the tree view and toolbar button.

It is a key part of keeping the UI intuitive and predictable.

---

[Pinning System Index](#pinning-system-index)

---

















## **IsTreeNodePinnable**

### **IsTreeNodePinnable Index**
- **What this method does**  
- **How the method works**  
- **Why this method matters**  
- **Back to Pinning System Index**

---

### **What this method does**

`IsTreeNodePinnable` determines whether a given tree node represents a folder that can be pinned. It validates the node, checks the underlying path, and ensures the folder is eligible for pinning based on existence and special‑folder rules.

---

```vb.net
Private Function IsTreeNodePinnable(node As TreeNode) As Boolean
    If node Is Nothing Then Return False

    Dim path As String = TryCast(node.Tag, String)
    If String.IsNullOrEmpty(path) Then Return False
    If Not Directory.Exists(path) Then Return False
    If IsSpecialFolder(path) Then Return False

    Return True
End Function

```

---

### **How the method works**

- Confirms the node itself is not `Nothing`.  
- Extracts the folder path stored in the node’s `Tag`.  
- Validates that the path is not empty and that the directory actually exists.  
- Rejects the folder if it is a special folder such as Documents, Desktop, or Downloads.  
- Returns `True` only when the folder is a real, existing, non‑special directory.

This ensures the tree view only offers pin/unpin options for folders that are truly pinnable.

---

### **Why this method matters**

`IsTreeNodePinnable` ensures:

- The tree view context menu never shows invalid pin options.  
- System folders remain protected from being pinned.  
- The pinning system behaves consistently across both the tree view and file list.  
- The UI remains predictable and aligned with Explorer‑style behavior.

It is a small but essential guardrail that keeps the pinning workflow safe and intuitive.

---

[Pinning System Index](#pinning-system-index)

---





















## **UpdateTreeContextMenu**

### **UpdateTreeContextMenu Index**
- **What this method does**  
- **How the method works**  
- **Why this method matters**  
- **Back to Pinning System Index**

---

### **What this method does**

`UpdateTreeContextMenu` updates the visibility of the **Pin** and **Unpin** options in the tree view’s context menu. It evaluates the selected node, validates the underlying folder path, and shows the correct action based on whether the folder is already pinned.

---

```vb.net
Private Sub UpdateTreeContextMenu(node As TreeNode)
    Dim path As String = TryCast(node.Tag, String)

    If String.IsNullOrEmpty(path) OrElse
   Not Directory.Exists(path) OrElse
   IsSpecialFolder(path) Then

        mnuPin.Visible = False
        mnuUnpin.Visible = False
        Exit Sub
    End If

    mnuPin.Visible = Not IsPinned(path)
    mnuUnpin.Visible = IsPinned(path)
End Sub


```

---

### **How the method works**

- Extracts the folder path from the selected tree node’s `Tag`.  
- Validates the path by checking that:  
  - It is not empty.  
  - The directory exists.  
  - It is not a special folder (Documents, Desktop, Downloads, etc.).  
- If the path is invalid, both **Pin** and **Unpin** are hidden to prevent incorrect actions.  
- If the path is valid, the method checks whether the folder is pinned:  
  - Shows **Pin** when the folder is not pinned.  
  - Shows **Unpin** when the folder is pinned.

This ensures the context menu always presents the correct, safe action for the selected folder.

---

### **Why this method matters**

`UpdateTreeContextMenu` ensures:

- The tree view behaves consistently with the file list and toolbar button.  
- Users never see pin options for invalid or protected folders.  
- The UI remains predictable and Explorer‑accurate.  
- The pinning workflow stays intuitive regardless of where the user interacts.

It is a key part of keeping the tree view’s behavior aligned with the rest of the pinning system.

---

[Pinning System Index](#pinning-system-index)

---




















## **IsSpecialFolder**

### **IsSpecialFolder Index**
- **What this method does**  
- **How the method works**  
- **Why this method matters**  
- **Back to Pinning System Index**

---

### **What this method does**

`IsSpecialFolder` determines whether a given path refers to a system‑managed folder that should never be pinnable. It compares the input path against a predefined set of known special folders such as Documents, Desktop, Downloads, and other user‑profile locations.

---

```vb.net
Private Function IsSpecialFolder(folderPath As String) As Boolean
    Dim specialFolders As (String, String)() = {
      ("Documents", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)),
      ("Music", Environment.GetFolderPath(Environment.SpecialFolder.MyMusic)),
      ("Pictures", Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)),
      ("Videos", Environment.GetFolderPath(Environment.SpecialFolder.MyVideos)),
      ("Downloads", IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads")),
      ("Desktop", Environment.GetFolderPath(Environment.SpecialFolder.Desktop))
    }

    Return specialFolders.Any(Function(sf) String.Equals(sf.Item2, folderPath, StringComparison.OrdinalIgnoreCase))
End Function


```

---

### **How the method works**

- Builds a list of special folder paths using `Environment.GetFolderPath` and known user‑profile subdirectories.  
- Normalizes both the input path and each special folder path for consistent comparison.  
- Checks for an exact match between the normalized input path and any special folder path.  
- Returns `True` when the folder is considered special; otherwise returns `False`.

This ensures the method reliably identifies folders that should not be pinned.

---

### **Why this method matters**

`IsSpecialFolder` protects:

- System‑managed directories that users should not modify through pinning.  
- Core user folders like Documents, Desktop, Downloads, Pictures, Music, and Videos.  
- The integrity of the pinning system by preventing unsafe or confusing pin states.  
- Consistency with Windows Explorer, which also restricts pinning of these locations.

It acts as a safety gate for every pin/unpin operation.

---

[Pinning System Index](#pinning-system-index)

---























## **PinFromFiles / UnpinFromFiles**

### **PinFromFiles / UnpinFromFiles Index**
- **What these handlers do**  
- **How they work**  
- **Why they matter**  
- **Back to Pinning System Index**

---

### **What these handlers do**

`PinFromFiles` and `UnpinFromFiles` are the file‑list context‑menu handlers that let the user pin or unpin a folder directly from the list view. They extract the selected folder and pass it to the central toggle logic so the UI and Easy Access list update consistently.

---

---

```vb.net
Private Sub PinFromFiles_Click(sender As Object, e As EventArgs)
    If lvFiles.SelectedItems.Count = 0 Then Exit Sub
    TryPinOrUnpin(TryCast(lvFiles.SelectedItems(0).Tag, String))
End Sub

Private Sub UnpinFromFiles_Click(sender As Object, e As EventArgs)
    If lvFiles.SelectedItems.Count = 0 Then Exit Sub
    TryPinOrUnpin(TryCast(lvFiles.SelectedItems(0).Tag, String))
End Sub

```

---

### **How they work**

- They first confirm that an item is selected in the file list.  
- They extract the folder path from the selected item’s `Tag`.  
- They call `TogglePin(path)` to perform the actual pin or unpin operation.  
- The toggle method then updates the Easy Access file and refreshes the UI.

These handlers contain no business logic themselves—they simply act as bridges between the UI and the core pinning engine.

---

### **Why they matter**

`PinFromFiles` and `UnpinFromFiles` ensure:

- The file list behaves consistently with the tree view and toolbar.  
- Users can pin or unpin folders using familiar right‑click actions.  
- All pin/unpin operations flow through the same centralized logic (`TogglePin`).  
- The UI stays synchronized regardless of where the action originates.

They complete the GUI side of the pinning workflow.

---

[Pinning System Index](#pinning-system-index)

---





















## **Pin_Click / Unpin_Click**

### **Pin_Click / Unpin_Click Index**
- **What these handlers do**  
- **How they work**  
- **Why they matter**  
- **Back to Pinning System Index**

---

### **What these handlers do**

`Pin_Click` and `Unpin_Click` are the tree‑view context‑menu handlers that let the user pin or unpin a folder directly from the folder tree. They mirror the behavior of the file‑list handlers but operate on the currently selected tree node instead of a list item.

---

```vb.net

    Private Sub Pin_Click(sender As Object, e As EventArgs)
        TryPinOrUnpin(TryCast(tvFolders.SelectedNode?.Tag, String))
    End Sub

    Private Sub Unpin_Click(sender As Object, e As EventArgs)
        TryPinOrUnpin(TryCast(tvFolders.SelectedNode?.Tag, String))
    End Sub

```

---

### **How they work**

- They extract the folder path from the selected tree node’s `Tag`.  
- They validate that the node exists and contains a usable path.  
- They call `TogglePin(path)` to perform the actual pin or unpin operation.  
- The toggle logic updates the Easy Access file and refreshes the UI.

These handlers contain no business logic themselves—they simply delegate to the central pinning engine.

---

### **Why they matter**

`Pin_Click` and `Unpin_Click` ensure:

- The tree view offers the same pin/unpin functionality as the file list.  
- Users can pin or unpin folders using familiar right‑click actions in the tree.  
- All pin/unpin operations flow through the unified `TogglePin` method.  
- The UI stays synchronized regardless of where the action originates.

They complete the tree‑view side of the pinning workflow.

---

[Pinning System Index](#pinning-system-index)

---























## **TogglePin**

### **TogglePin Index**
- [What this method does](#what-this-method-does)  
- [How the method works](#how-the-method-works)  
- [Why this method matters](#why-this-method-matters)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`TogglePin` is the core method that switches a folder between pinned and unpinned states. It validates the folder, determines whether it is already pinned, performs the appropriate action, and then refreshes the UI so the change is immediately visible.

---

```vb.net
Private Sub TogglePin(path As String)
    If String.IsNullOrWhiteSpace(path) Then Exit Sub
    If Not Directory.Exists(path) Then Exit Sub
    If IsSpecialFolder(path) Then Exit Sub

    Dim name As String = GetFolderDisplayName(path)

    If IsPinned(path) Then
        RemoveFromEasyAccess(path)
    Else
        AddToEasyAccess(name, path)
    End If

    RefreshPinUI()
End Sub
```

---

### **How the method works**

- It first checks whether the provided path is empty or whitespace. If so, the method exits immediately.  
- It verifies that the path points to an existing directory. If not, the operation is ignored.  
- It checks whether the folder is a special folder (Documents, Desktop, etc.). Special folders cannot be pinned, so the method exits.  
- It retrieves a display name for the folder using `GetFolderDisplayName(path)`.  
- It calls `IsPinned(path)` to determine the current state:  
  - If the folder **is pinned**, it calls `RemoveFromEasyAccess(path)` to unpin it.  
  - If the folder **is not pinned**, it calls `AddToEasyAccess(name, path)` to pin it.  
- Finally, it calls `RefreshPinUI()` to update the tree view, file list, and pin button.

---

### **Why this method matters**

`TogglePin` is the heart of the pinning system. It ensures:

- The pin/unpin behavior is consistent everywhere (CLI, tree view, file list, toolbar).  
- Invalid or unsafe paths are rejected early.  
- The UI always reflects the correct state after any change.  
- The logic for pinning is centralized, preventing duplication across the app.

---

[Pinning System Index](#pinning-system-index)

---





















## **GetPinnableTarget**

### **GetPinnableTarget Index**
- [What this method does](#what-this-method-does-2)  
- [How the method works](#how-the-method-works-2)  
- [Why this method matters](#why-this-method-matters-2)  
- [Back to Pinning System Index](#pinning-system-index)

---

### **What this method does**

`GetPinnableTarget` determines which folder should be pinned or unpinned based on the user’s most recent interaction. It inspects the last‑focused control and returns the appropriate folder path, or `Nothing` if no valid target exists.

---

```vb.net
Private Function GetPinnableTarget() As String

    ' ==========================
    ' 0. Helper: must be a real, pinnable directory
    ' ==========================
    Dim isValidDir As Func(Of String, Boolean) =
    Function(p As String)
        Return Not String.IsNullOrEmpty(p) AndAlso
               Directory.Exists(p) AndAlso
               Not IsSpecialFolder(p)
    End Function

    ' ==========================
    ' 1. If the last focused control was lvFiles
    ' ==========================
    If _lastFocusedControl Is lvFiles AndAlso lvFiles.SelectedItems.Count > 0 Then
        Dim path As String = TryCast(lvFiles.SelectedItems(0).Tag, String)
        If isValidDir(path) Then Return path
    End If

    ' ==========================
    ' 2. If the last focused control was tvFolders
    ' ==========================
    If _lastFocusedControl Is tvFolders AndAlso tvFolders.SelectedNode IsNot Nothing Then
        Dim path As String = TryCast(tvFolders.SelectedNode.Tag, String)
        If isValidDir(path) Then Return path
    End If

    ' ==========================
    ' 3. If the last focused control was the address bar
    ' ==========================
    If _lastFocusedControl Is txtAddressBar Then
        If isValidDir(currentFolder) Then Return currentFolder
    End If

    ' ==========================
    ' 4. No contextual target
    ' ==========================
    Return Nothing
End Function

```

---

### **How the method works**

The method evaluates potential targets in a strict priority order:

1. **File list**  
   If the last focused control was the file list and a folder is selected, that folder is returned.

2. **Tree view**  
   If the last focused control was the tree view and a folder is selected, that folder is returned.

3. **Address bar**  
   If the last focused control was the address bar and the current folder is valid, the current folder is returned.

4. **Fallback**  
   If none of the above conditions apply, the method returns `Nothing`.

This ensures the pin button and the `pin` command always act on the folder the user is actually interacting with.

---

### **Why this method matters**

`GetPinnableTarget` is essential for:

- Making the pin/unpin button context‑aware  
- Ensuring the `pin` command behaves predictably  
- Preventing accidental pinning of the wrong folder  
- Keeping keyboard, mouse, and CLI interactions unified  
- Supporting a consistent mental model across the entire app  

Without this method, the pinning system would not know which folder the user intends to pin, especially when switching between the tree view, file list, and address bar.

---

[Pinning System Index](#pinning-system-index)

---

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






[ Table of Contents](#table-of-contents)









---
---
---
---










# `Find Command`







<img width="1266" height="713" alt="108" src="https://github.com/user-attachments/assets/042a5223-4f0a-4926-9068-76e7f4075c78" />







**This code takes the user’s search term, finds matching files in the current folder, highlights the first match, shows how many results there are, and gives the user a clean, Explorer‑style search experience.**

Think of this code as the “start a search” feature in our file explorer.  
When the user types something like:


```

find report

```

this block of code runs.


```vb

            Case "find", "search"
                If parts.Length > 1 Then

                    Dim searchTerm As String = String.Join(" ", parts.Skip(1)).Trim()

                    If String.IsNullOrWhiteSpace(searchTerm) Then
                        ShowStatus(
                            StatusPad & IconDialog &
                            "  Usage: find [search_term]   Example: find document"
                        )
                        Return
                    End If

                    ' Announce search
                    ShowStatus(StatusPad & IconSearch & "  Searching for: " & searchTerm)

                    ' Perform search
                    OnlySearchForFilesInCurrentFolder(searchTerm)

                    ' Reset index for new search
                    SearchIndex = 0
                    RestoreBackground()

                    ' If results exist, auto-select the first one
                    If SearchResults.Count > 0 Then
                        lvFiles.SelectedItems.Clear()
                        SelectListViewItemByPath(SearchResults(0))

                        Dim nextPath As String = SearchResults(SearchIndex)
                        Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

                        lvFiles.Focus()
                        HighlightSearchMatches()

                        ' Unified search HUD
                        ShowStatus(
                        StatusPad & IconSearch &
                        $"  Result {SearchIndex + 1} of {SearchResults.Count}    ""{fileName}""    Next  F3    Open  Ctrl+O    Reset  Esc"
                    )
                    Else
                        ShowStatus(
                        StatusPad & IconDialog &
                        "  No results found for: " & searchTerm
                    )
                    End If

                Else
                    ShowStatus(
                        StatusPad & IconDialog &
                        "  Usage: find [search_term]   Example: find document"
                    )
                End If

                Return

```


Let’s go through it step by step.


 **Make sure the user actually typed a search term**

```vb
If parts.Length > 1 Then
```

- `parts` is the command split into words.
- If the user typed only `find` with nothing after it, then there’s nothing to search for.


 **Combine everything after the command into one search term**

```vb
Dim searchTerm As String = String.Join(" ", parts.Skip(1)).Trim()
```

- This takes everything after the word `find` and turns it into one string.
- Example:  
  `find my report` → `"my report"`


 **If the search term is empty, show a helpful message**

```vb
If String.IsNullOrWhiteSpace(searchTerm) Then
    ShowStatus("Usage: find [search_term] Example: find document")
    Return
End If
```

- This protects the user from mistakes.
- It stops the command early and explains how to use it.


 **Tell the user that the search has started**

```vb
ShowStatus("Searching for: " & searchTerm)
```

- This updates your status bar so the user knows something is happening.


 **Actually perform the search**

```vb
OnlySearchForFilesInCurrentFolder(searchTerm)
```

- This is your search engine.
- It looks through the current folder and fills `SearchResults` with matching file paths.


 **Reset the search index**

```vb
SearchIndex = 0
RestoreBackground()
```

- Since this is a *new* search, you always start at the first result.
- `RestoreBackground()` clears any old highlighting from a previous search.


 **If matches were found…**

```vb
If SearchResults.Count > 0 Then
```

Now the fun part begins.


 **Select the first matching file**

```vb
lvFiles.SelectedItems.Clear()
SelectListViewItemByPath(SearchResults(0))
```

- Clears any previous selection.
- Highlights the first matching file in the ListView.
- Scrolls to it so the user can see it.


 **Get a clean file name for display**

```vb
Dim nextPath As String = SearchResults(SearchIndex)
Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)
```

- Turns something like  
  `C:\Users\Joseph\Documents\Report1.pdf`  
  into  
  `"Report1"`


 **Focus the ListView and highlight all matches**

```vb
lvFiles.Focus()
HighlightSearchMatches()
```

- Gives keyboard focus to the file list.
- Highlights all matching items.


 **Show the unified search HUD**

```vb
ShowStatus(
    $"Result {SearchIndex + 1} of {SearchResults.Count}  ""{fileName}""  Next F3  Open Ctrl+O  Reset Esc"
)
```

This status message that tells the user:

- Which result they’re on  
- How many results exist  
- The file name  
- Helpful keyboard shortcuts  


 **If no results were found…**

```vb
Else
    ShowStatus("No results found for: " & searchTerm)
End If
```

- The app gently tells the user nothing matched.


 **If the user typed only “find” with no term**

```vb
Else
    ShowStatus("Usage: find [search_term] Example: find document")
End If
```

- Another friendly reminder about how to use the command.





<img width="1266" height="713" alt="106" src="https://github.com/user-attachments/assets/28ad95b5-5ccb-4365-bfae-a611108a07f9" />












[ Table of Contents](#table-of-contents)


---
---
---
---

































# `HandleFindNextCommand`

<img width="1266" height="713" alt="106" src="https://github.com/user-attachments/assets/28ad95b5-5ccb-4365-bfae-a611108a07f9" />

This method cycles to the next search result, highlights it in the file list, scrolls to it, and updates the status bar so the user always knows where they are and what they can do next.

Below is the full code, then we’ll walk through it one step at a time.


```vb.net

    Private Sub HandleFindNextCommand()

        ' No active search
        If SearchResults.Count = 0 Then
            ShowStatus(
            StatusPad & IconDialog &
            "  No previous search results. Press Ctrl+F or enter: find [search_term] to start a search."
        )
            Return
        End If

        ' Advance index with wraparound
        SearchIndex += 1
        If SearchIndex >= SearchResults.Count Then
            SearchIndex = 0
        End If

        ' Select the next result
        lvFiles.SelectedItems.Clear()
        Dim nextPath As String = SearchResults(SearchIndex)
        SelectListViewItemByPath(nextPath)

        Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)

        ' Status HUD
        ShowStatus(
            StatusPad & IconSearch &
            $"  Result {SearchIndex + 1} of {SearchResults.Count}    ""{fileName}""    Next  F3    Open  Ctrl+O    Reset  Esc"
        )
    End Sub


```


```vb
Private Sub HandleFindNextCommand()
```
- This defines a **subroutine** (a block of code) named `HandleFindNextCommand`.
- It runs whenever the user triggers the “Find Next” action — for example, pressing **F3** or typing `findnext`.

---

```vb
' No active search
If SearchResults.Count = 0 Then
```
- This checks whether there are **any search results stored**.
- `SearchResults` is a list of file paths found by the last search.
- If the list is empty (`Count = 0`), it means the user hasn’t searched yet.

---

```vb
    ShowStatus(
        StatusPad & IconDialog &
        "  No previous search results. Press Ctrl+F or enter: find [search_term] to start a search."
    )
```
- This displays a friendly message in your status bar.
- `StatusPad` and `IconDialog` are just decorative pieces that make the message look nice.
- The message tells the user how to start a search.

---

```vb
    Return
End If
```
- `Return` stops the subroutine immediately.
- Since there are no results, there’s nothing else to do.

---

```vb
' Advance index with wraparound
SearchIndex += 1
```
- This moves to the **next** search result.
- `SearchIndex` keeps track of which result is currently selected.
- `+= 1` means “add 1 to the current value.”

---

```vb
If SearchIndex >= SearchResults.Count Then
    SearchIndex = 0
End If
```
- This handles **wraparound**.
- If the index goes past the last result, it jumps back to the first one.
- This creates a loop: last → first → second → etc.
- It feels like how Explorer cycles through search results.

---

```vb
' Select the next result
lvFiles.SelectedItems.Clear()
```
- Clears any previous selection in the file list.
- This ensures only the new result is highlighted.

---

```vb
Dim nextPath As String = SearchResults(SearchIndex)
```
- Retrieves the file path of the next search result.
- Stores it in a variable named `nextPath`.

---

```vb
SelectListViewItemByPath(nextPath)
```
- This is our helper function that:
  - Finds the matching item in the ListView.
  - Selects the item.
  - Scrolls the ListView so the item is visible.
- This is what makes the UI jump to the matching file.

---

```vb
Dim fileName As String = Path.GetFileNameWithoutExtension(nextPath)
```
- Extracts just the **file name** (no folder, no extension).
- This is used for a cleaner status message.

---

```vb
' Status HUD
ShowStatus(
    StatusPad & IconSearch &
    $"  Result {SearchIndex + 1} of {SearchResults.Count}    ""{fileName}""    Next  F3    Open  Ctrl+O    Reset  Esc"
)
```
- Updates the status bar with a compact “search HUD.”
- It shows:
  - Which result you’re on  
    (`Result 3 of 12`)
  - The file name  
    (`"Report Q1"`)
  - Helpful keyboard shortcuts  
    (`Next F3`, `Open Ctrl+O`, `Reset Esc`)
- `SearchIndex + 1` is used because lists start at 0, but humans start at 1.

---

```vb
End Sub
```
- Marks the end of the subroutine.











<img width="1266" height="713" alt="107" src="https://github.com/user-attachments/assets/220aeaa9-4150-4e47-85bb-467c86df8b79" />



[ Table of Contents](#table-of-contents)


---
---
---
---






# `SelectListViewItemByPath`




<img width="1266" height="713" alt="106" src="https://github.com/user-attachments/assets/28ad95b5-5ccb-4365-bfae-a611108a07f9" />





This helper is a quiet hero in our search system. It ensures that:

- The correct file is highlighted  
- The UI scrolls to it  
- Keyboard navigation starts from the right place  
- The user feels like the app “knows where they are”

Below is the full code, then we’ll walk through it one step at a time.

```vb.net

    Private Sub SelectListViewItemByPath(fullPath As String)
        For Each item As ListViewItem In lvFiles.Items
            If String.Equals(item.Tag.ToString(), fullPath, StringComparison.OrdinalIgnoreCase) Then
                item.Selected = True
                item.Focused = True
                item.EnsureVisible()
                Exit For
            End If
        Next
    End Sub

```


```vb.net
Private Sub SelectListViewItemByPath(fullPath As String)
```
- This defines a **subroutine** named `SelectListViewItemByPath`.
- It takes one input: `fullPath`, which is the complete file or folder path you want to select in the ListView.
- The goal is simple: find the matching item and highlight it.


```vb
For Each item As ListViewItem In lvFiles.Items
```
- This starts a loop that goes through **every item** in the ListView (`lvFiles`).
- `item` represents the current ListViewItem being checked.
- The loop continues until a match is found or all items have been checked.


```vb
    If String.Equals(item.Tag.ToString(), fullPath, StringComparison.OrdinalIgnoreCase) Then
```
- Each ListViewItem stores its full path inside its `.Tag` property.
- This line compares the item's stored path with the `fullPath` we’re searching for.
- `StringComparison.OrdinalIgnoreCase` means:
  - Compare the text exactly
  - Ignore uppercase/lowercase differences  
  This matches how Windows treats file paths.


```vb
        item.Selected = True
```
- Marks the item as **selected** in the ListView.
- This highlights it visually.


```vb
        item.Focused = True
```
- Gives the item **keyboard focus**.
- This ensures:
  - Arrow keys start from this item
  - The dotted focus rectangle appears
  - It behaves like the user clicked it


```vb
        item.EnsureVisible()
```
- Scrolls the ListView so the item is visible.
- If the item is off‑screen, the ListView automatically scrolls to bring it into view.
- This is essential for a smooth “Find Next” experience.


```vb
        Exit For
```
- Stops the loop immediately.
- No need to keep searching once the correct item is found.


```vb
    End If
Next
```
- Ends the `If` block and continues the loop if needed.


```vb
End Sub
```
- Ends the subroutine.





[ Table of Contents](#table-of-contents)













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





<img width="1920" height="1080" alt="120" src="https://github.com/user-attachments/assets/8a2b6f6e-9d53-401d-bf46-d3104378d6f9" />













<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/af1e3f4e-e007-4679-af4a-945de7a8fa78" />











<img width="1920" height="1080" alt="109" src="https://github.com/user-attachments/assets/452552a8-38b6-4296-bcd2-3183cc2c6b72" />













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






<img width="1920" height="1080" alt="116" src="https://github.com/user-attachments/assets/576fd1cf-93df-4fc6-af4a-a771dda9c8f1" />












<img width="1920" height="1080" alt="112" src="https://github.com/user-attachments/assets/026bb51d-6cb4-43b6-b17e-4dd2e43134f7" />












<img width="1920" height="1080" alt="111" src="https://github.com/user-attachments/assets/3fa71463-e23d-42de-98ee-dee4cce7f1c0" />



<img width="1920" height="1080" alt="110" src="https://github.com/user-attachments/assets/17878ab7-5ee8-443b-b455-7c9e5c462d91" />















<img width="1920" height="1080" alt="090" src="https://github.com/user-attachments/assets/056425b3-45e9-4c3d-a14c-d1104bf20399" />


<img width="1920" height="1080" alt="011" src="https://github.com/user-attachments/assets/0b1bff67-26a9-456d-9c42-0b96d4dd1bb8" />


[ Table of Contents](#table-of-contents)

---
---
---
---


