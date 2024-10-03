
# Excel OpenAI Processor

A streamlined Python script that processes Excel files by reading vulnerability data and leveraging OpenAI's API to categorize risks, assess exploitability, and generate impact descriptions. This tool is designed for cybersecurity professionals to automate and enhance vulnerability assessments efficiently.

## Table of Contents
- Features
- Prerequisites
- Installation
- Configuration
- Usage
- Project Structure
- Troubleshooting
- Contributing
- License
- Contact

## Features
- **Automated Risk Categorization**: Classifies vulnerabilities into Confidentiality, Integrity, or Availability based on the CIA triad.
- **Exploitability Assessment**: Evaluates how easy or difficult it is to exploit each vulnerability, categorizing them as Low, Medium, or High.
- **Impact Description Generation**: Automatically generates detailed impact descriptions for vulnerabilities lacking this information.
- **Excel Integration**: Reads from and writes to Excel files, ensuring seamless data handling.
- **Environment Variable Management**: Utilizes global environment variables for secure and efficient API key management.
- **Logging**: Provides informative logs to monitor processing status and debug issues.
- **Simplified Script**: Easy-to-understand and modify script suitable for both beginners and professionals.

## Prerequisites

Before setting up the Excel OpenAI Processor, ensure you have the following:

### Python Installed:
- Python 3.6 or later is required.
- Download and install Python from [python.org](https://www.python.org/).

### Git Installed:
- Required for cloning the repository.
- Download and install Git from [git-scm.com](https://git-scm.com/).

### OpenAI API Key:
- Obtain your API key from [OpenAI's API Keys](https://beta.openai.com/account/api-keys).
- Ensure you have access to the GPT-4 model.

### Excel File:
Prepare your input Excel file (`vulnerabilities.xlsx`) with the following headers:
- ID
- Vulnerability
- Risk Category
- Severity
- Exploitability
- Description
- Impact
- Remediation
- References

## Installation

Follow these steps to set up the project on your local machine.

1. **Clone the Repository**

   Open Command Prompt or PowerShell and navigate to your desired directory. Then, clone the repository:

   ```bash
   git clone https://github.com/yourusername/excel-openai-processor.git
   ```

   Replace `yourusername` with your actual GitHub username.

2. **Navigate to the Project Directory**

   ```bash
   cd excel-openai-processor
   ```

3. **Create a Virtual Environment**

   It's recommended to use a virtual environment to manage dependencies.

   ```bash
   python -m venv venv
   ```

4. **Activate the Virtual Environment**

   - Command Prompt:
   
     ```cmd
     venv\Scripts\activate.bat
     ```

   - PowerShell:
   
     ```powershell
     venv\Scripts\Activate.ps1
     ```

   If you encounter an execution policy error in PowerShell, run the following command to allow script execution for the current session:

   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process
   ```

5. **Install Dependencies**

   ```bash
   pip install pandas openpyxl openai
   ```

## Configuration

1. **Set Global Environment Variable for OpenAI API Key**

   To securely manage your OpenAI API key, set it as a global environment variable in Windows.

   a. **Open System Properties**
   - Press `Win + Pause/Break` keys simultaneously.
   - Alternatively, right-click on `This PC` on the desktop or in File Explorer and select `Properties`, then click `Advanced system settings`.

   b. **Access Environment Variables**
   - In the System Properties window, navigate to the `Advanced` tab.
   - Click on the `Environment Variables...` button at the bottom.

   c. **Add a New System Variable**
   - Under the System variables section, click `New...`.
   - Variable name: `OPENAI_API_KEY`
   - Variable value: `your-openai-api-key-here`
   - Click OK to save.

   Replace `your-openai-api-key-here` with your actual OpenAI API key.

   d. **Verify the Variable**

   Open PowerShell or Command Prompt.

   Run the following command:

   ```powershell
   echo $env:OPENAI_API_KEY
   ```

   You should see your API key displayed.

2. **Prepare the Excel File**

   Ensure your input Excel file (`vulnerabilities.xlsx`) is placed in the project directory and structured with the required headers.

## Usage

Run the Python script to process the Excel file.

1. **Activate the Virtual Environment (If Not Already Active)**

   - Command Prompt:
   
     ```cmd
     venv\Scripts\activate.bat
     ```

   - PowerShell:
   
     ```powershell
     venv\Scripts\Activate.ps1
     ```

2. **Execute the Script**

   ```bash
   python process_excel.py
   ```

3. **Monitor the Output**

   The script will log processing steps in the console.
   Upon completion, an updated Excel file (`vulnerabilities_updated.xlsx`) will be generated in the project directory with the following updates:
   - **Risk Category**: Categorized as Confidentiality, Integrity, or Availability.
   - **Exploitability**: Assessed as Low, Medium, or High.
   - **Impact**: Generated if previously empty.

## Project Structure

```bash
excel-openai-processor/
├── venv/                        # Virtual environment directory
├── vulnerabilities.xlsx         # Input Excel file
├── vulnerabilities_updated.xlsx # Output Excel file
├── process_excel.py             # Main Python script
├── .gitignore                   # Specifies files to ignore in Git
├── README.md                    # Project documentation
└── requirements.txt             # Project dependencies (optional)
```

Note: The `venv/` directory can be excluded from version control.

## Troubleshooting

1. **OpenAI API Key Not Found**

   - **Symptom**: Error message indicating that the OpenAI API key is missing.
   - **Solution**:
     - Ensure that the `OPENAI_API_KEY` environment variable is correctly set.
     - Open a new terminal session after setting the environment variable.
     - Verify by running:
     ```powershell
     echo $env:OPENAI_API_KEY
     ```

2. **Excel File Not Found**

   - **Symptom**: Error indicating that the input Excel file is missing.
   - **Solution**:
     - Ensure that `vulnerabilities.xlsx` is placed in the project directory.
     - Verify the file name and path.

3. **Permission Issues**

   - **Symptom**: Errors related to file access permissions.
   - **Solution**:
     - Run the terminal or PowerShell as an administrator.
     - Ensure that you have read/write permissions for the project directory and Excel files.

4. **Rate Limiting by OpenAI**

   - **Symptom**: Errors indicating too many requests to the OpenAI API.
   - **Solution**:
     - Increase the `time.sleep(1)` delay between API requests in the script.
     - Monitor your OpenAI API usage and adjust based on your plan.

5. **Missing Dependencies**

   - **Symptom**: Import errors for missing Python packages.
   - **Solution**:
     - Ensure all dependencies are installed using:
     ```bash
     pip install pandas openpyxl openai
     ```

## Contributing

Contributions are welcome! To contribute:

1. **Fork the Repository**
2. **Create a New Branch**

   ```bash
   git checkout -b feature/YourFeatureName
   ```

3. **Commit Your Changes**

   ```bash
   git commit -m "Add Your Feature"
   ```

4. **Push to the Branch**

   ```bash
   git push origin feature/YourFeatureName
   ```

5. **Open a Pull Request**

   Provide a clear description of your changes and the problem they address.

## License

This project is licensed under the MIT License.

## Contact

For any questions or support, please contact `saad.alaref95@gmail.com`.
