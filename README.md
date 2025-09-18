Converter-Ultimate
Your all-in-one solution for intelligent file conversion. Built with Python, this powerful utility features a modern interface, support for a wide range of formats, and optional AI-powered features with GPU acceleration. This project is fully open-source.

Installation and Usage
This guide provides instructions for Windows and Linux users.

For Windows Users
The recommended method for Windows is to use the runtime.py smart installer, which handles all dependencies and builds the application for you.

1. Install Python
If you don't have Python, download and install the latest version from the official website.

Go to: python.org

Important: During installation, make sure to check the box that says Add Python to PATH.

2. Install the GUI Prerequisite
The installer needs a library to display its window. Open a Command Prompt (cmd.exe) and run this command:

pip install PyQt6

3. Download and Run the Installer

Download the runtime.py script from the Releases Page of this repository.

Save it to a convenient location, like your Desktop.

Open a Command Prompt in that location and run the script:

python runtime.py

4. Wait for Installation
The installer window will appear and will automatically download all required libraries, build the final .exe, and create shortcuts. This may take several minutes.

5. Launch the Application
Once the installation is complete, you can find and run the application by searching for Converter-Ultimate in your Start Menu.

For Linux Users
On Linux, you will set up and run the application source code directly.

1. Clone the Repository
Open your terminal and clone the project:

git clone [https://github.com/SomerandmguyintheInternet/Converter-ultimate.git](https://github.com/SomerandmguyintheInternet/Converter-ultimate.git)
cd Converter-ultimate

2. Create a Virtual Environment
It is highly recommended to use a virtual environment.

python3 -m venv .venv
source .venv/bin/activate

3. Install Dependencies
Install all required libraries using the requirements.txt file:

pip install -r requirements.txt

4. Run the Application
Launch the main application script directly:

python app.py

System Requirements
To ensure a smooth experience, please review the minimum and recommended system specifications.

Minimum Requirements
Operating System: Windows 10 (64-bit) or a modern Linux distribution

CPU: 2 Cores / 4 Threads @ 2.5 GHz+

RAM: 8 GB

Notes: Internet access is required for the installer to download dependencies.

Recommended Specifications
Operating System: Windows 11 (64-bit) or a modern Linux distribution

CPU: 4 Cores / 8 Threads @ 3.5 GHz+ (Turbo)

RAM: 16 GB

Notes: A dedicated NVIDIA GPU is recommended for GPU-accelerated tasks.

Application Features
Multi-Format Conversion: Effortlessly convert between various document, image, and data formats.

AI-Powered Features: (Optional) Leverage state-of-the-art AI for tasks like document summarization and analysis.

GPU Acceleration: (Optional) Enable GPU support via PyTorch to dramatically speed up complex conversion and AI tasks.

Modular Installation: Choose which optional features you want to install.

Developer Console: (Optional) An integrated console for advanced users and debugging.

Modern UI: A clean, intuitive, and easy-to-use interface.

Self-Contained: The installer packages the final application into a single .exe with all necessary dependencies.

Project Internals
This repository uses a two-script system to create a robust and seamless user experience on Windows.

runtime.py (The Smart Installer)
This script is the user-facing installer. Its job is not to be the application, but to build and deploy it. It handles dependency checking, downloads the latest application source, packages it into a final .exe using PyInstaller, and creates all necessary system shortcuts. It also serves as the uninstaller.

app.py (The Core Application)
This is the actual file conversion tool. The installer downloads this file and compiles it into the final Converter-Ultimate.exe that users interact with. It contains all the UI, conversion logic, and feature integrations.
