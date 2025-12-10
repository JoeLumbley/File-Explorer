# 📁 File Explorer

**File Explorer** is a simple and user-friendly file management application designed to provide an intuitive interface for navigating, copying, moving, and deleting files and directories.



<img width="1920" height="1080" alt="009" src="https://github.com/user-attachments/assets/dcdc025a-ba88-46a9-8247-df93db2da420" />


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

- **Change Directory (`cd`)**: 
  - Usage: `cd [directory]`
  - Changes the current working directory to the specified path.
  - Example: `cd C:\Users\Documents`

- **Copy Files (`copy`)**: 
  - Usage: `copy [source] [destination]`
  - Copies files or directories from a source to a destination.
  - Example: `copy C:\folder1\file.txt C:\folder2`

<img width="1920" height="1080" alt="010" src="https://github.com/user-attachments/assets/909d297d-8e29-4b8c-a07e-e8301898aae5" />


- **Move Files (`move`)**: 
  - Usage: `move [source] [destination]`
  - Moves files or directories from one location to another.
  - Example: `move C:\folder1\file.txt C:\folder2\file.txt`

- **Delete Files (`delete`)**: 
  - Usage: `delete [file_or_directory]`
  - Deletes the specified file or directory.
  - Example: `delete C:\folder1\file.txt`

- **Help (`help`)**: 
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












<img width="1920" height="1080" alt="011" src="https://github.com/user-attachments/assets/0b1bff67-26a9-456d-9c42-0b96d4dd1bb8" />




