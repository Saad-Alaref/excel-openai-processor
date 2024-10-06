
# Excel Processor with OpenAI Integration

This project provides a Python-based tool for processing data from Excel spreadsheets. The tool is purpose-agnostic and easily configurable, making it applicable to a wide range of use cases, such as data enrichment, analysis and automation using OpenAI's GPT models.

## Features

- **Excel Processing**: Load and process rows from an Excel file.
- **OpenAI Integration**: Automatically generate content for specified columns using OpenAI's API.
- **Retry Logic**: Configurable retry mechanism for handling API calls.
- **Logging**: Detailed logging to monitor progress and troubleshoot errors.
- **Flexible Output**: Supports free-form text output or function-based structured responses from OpenAI.

## Requirements

- Python 3.8+
- Packages:
  - `pandas`
  - `openai`
  - `openpyxl`
  - `PyYAML`

## Setup

1. Clone the repository:
   ```bash
   git clone https://github.com/Saad-Alaref/excel-openai-processor.git
   ```
   ```bash
   cd excel-openai-processor
   ```

2. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

3. Set up your configuration file:
   - A sample configuration file is located in `config/config.yaml`. Modify it to suit your needs, such as specifying the input Excel file, OpenAI API key, and processing options.

4. Set the OpenAI API key as an environment variable:
   ```bash
   export OPENAI_API_KEY=your_openai_api_key
   ```

## Configuration

The tool uses a YAML configuration file (`config/config.yaml`) that specifies the input Excel file, columns to process, and OpenAI settings.

Example `config.yaml`:
```yaml
excel:
  input_path: "data/input.xlsx"
  sheet_name: "Sheet1"

columns:
  input:
    - ColumnA
    - ColumnB
  output:
    ColumnC:
      prompt: "Please summarize the following information:"
      max_tokens: 50
      temperature: 0.7
      fetch_all: false

openai:
  api_key_env_var: "OPENAI_API_KEY"
  model: "gpt-4"
  system_message: "You are a helpful assistant."

processing:
  sleep_time: 1
  retry_attempts: 3
  retry_delay: 5
```

### Configuration Details:

- **Excel Section**:
  - `input_path`: Path to the Excel file.
  - `sheet_name`: Name of the sheet to process.
- **Columns Section**:
  - `input`: Columns from the Excel sheet used to generate the prompt, these columns are passed (per row) to the API along with the prompt.
  - `output`: Columns to fill based on the API's response.
    - `prompt`: Template for the API prompt.
    - `max_tokens`: Maximum token count for the API response.
    - `temperature`: Controls randomness of the API's output.
    - `fetch_all`: Whether to overwrite existing data or only fetch missing data.
- **OpenAI Section**:
  - `api_key_env_var`: Environment variable that holds the API key.
  - `model`: OpenAI model to use (e.g., GPT-4).
  - `system_message`: System message to provide context to the model.
- **Processing Section**:
  - `sleep_time`: Time to wait between API calls.
  - `retry_attempts`: Number of times to retry if an API call fails.
  - `retry_delay`: Time to wait between retries.

## Usage

Run the script using:

```bash
python scripts/process_excel.py
```

This will start processing the Excel file as per the configuration, making API requests to OpenAI for generating outputs, and filling the results back into the Excel file.

## Logging

Logs are saved in `log_file.log`. You can monitor the progress of the script, including API calls and row processing details.

## Customization

This tool is designed to be flexible and purpose-agnostic, allowing you to adapt it to various domains. Whether you're working with business data, scientific information, or other fields, you can customize the prompts and column mappings to suit your use case.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Contact

For any reason, you may contact me using the following email address: `saad.alaref95@gmail.com`.
