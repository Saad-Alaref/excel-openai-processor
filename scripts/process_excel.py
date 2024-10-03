# scripts/process_excel.py

import os
import time
import logging
import yaml
import pandas as pd
import openai
from dotenv import load_dotenv
from pathlib import Path
from typing import Any, Dict, Optional
from openai import OpenAI  # Updated import
from openpyxl import load_workbook

# Load environment variables from .env if available
load_dotenv()

# ---------------------------- Logging Configuration ---------------------------- #

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)

# ---------------------------- Configuration Loader ---------------------------- #

def load_config(config_path: str) -> Dict[str, Any]:
    """
    Loads the YAML configuration file.

    Args:
        config_path (str): Path to the YAML config file.

    Returns:
        Dict[str, Any]: Configuration parameters.
    """
    try:
        with open(config_path, 'r') as file:
            config = yaml.safe_load(file)
        logger.info(f"Configuration loaded from {config_path}")
        return config
    except Exception as e:
        logger.error(f"Failed to load configuration file: {e}")
        raise

# ---------------------------- OpenAI API Client ---------------------------- #

class OpenAIClient:
    def __init__(self, api_key: str, model: str, system_message: str):
        """
        Initializes the OpenAI client with the provided API key, model, and system message.

        Args:
            api_key (str): OpenAI API key.
            model (str): OpenAI model to use (e.g., "gpt-4", "gpt-3.5-turbo").
            system_message (str): The system message to set the context for the AI assistant.
        """
        self.client = OpenAI(api_key=api_key)  # Initialize OpenAI client
        self.model = model
        self.system_message = system_message
        logger.info(f"OpenAI client initialized with model '{self.model}' and system message: '{self.system_message}'.")

    def create_completion(self, prompt: str, max_tokens: int, temperature: float) -> Optional[str]:
        """
        Sends a request to the OpenAI API to generate a completion.

        Args:
            prompt (str): The prompt to send to the API.
            max_tokens (int): Maximum number of tokens in the response.
            temperature (float): Sampling temperature.

        Returns:
            Optional[str]: The generated text or None if failed.
        """
        try:
            completion = self.client.chat.completions.create(
                model=self.model,
                messages=[
                    {"role": "system", "content": self.system_message},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature
            )
            # Corrected access to 'content'
            content = completion.choices[0].message.content.strip()
            logger.debug(f"OpenAI response: {content}")
            return content
        except Exception as e:
            logger.error(f"OpenAI API request failed: {e}")
            return None

# ---------------------------- Excel Processor ---------------------------- #

class ExcelProcessor:
    def __init__(self, config: Dict[str, Any], openai_client: OpenAIClient):
        """
        Initializes the Excel processor with configuration and OpenAI client.

        Args:
            config (Dict[str, Any]): Configuration parameters.
            openai_client (OpenAIClient): Instance of OpenAIClient.
        """
        self.config = config
        self.openai = openai_client
        self.input_path = config['excel']['input_path']
        self.sheet_name = config['excel'].get('sheet_name', 'Sheet1')
        self.input_columns = config['columns']['input']
        self.output_columns = config['columns']['output']
        self.sleep_time = config['processing'].get('sleep_time', 1)
        self.retry_attempts = config['processing'].get('retry_attempts', 3)
        self.retry_delay = config['processing'].get('retry_delay', 5)

        # Load the workbook once to keep it open during processing
        try:
            self.workbook = load_workbook(filename=self.input_path)
            self.sheet = self.workbook[self.sheet_name]
            logger.info(f"Workbook '{self.input_path}' loaded successfully.")
        except Exception as e:
            logger.error(f"Failed to load workbook '{self.input_path}': {e}")
            raise

    def save_workbook(self):
        """
        Saves the Excel workbook.

        Raises:
            Exception: If saving fails.
        """
        try:
            self.workbook.save(self.input_path)
            logger.info(f"Workbook '{self.input_path}' saved successfully.")
        except Exception as e:
            logger.error(f"Failed to save workbook '{self.input_path}': {e}")
            raise

    def process_row(self, row_number: int, row_data: pd.Series):
        """
        Processes a single row by generating outputs using OpenAI and writing directly to the Excel cell.

        Args:
            row_number (int): The row number in the Excel sheet (1-based index).
            row_data (pd.Series): The DataFrame row data.
        """
        logger.info(f"Processing row {row_number}: {row_data[self.input_columns].to_dict()}")
        for output_column, output_config in self.output_columns.items():
            prompt_template = output_config['prompt']
            max_tokens = output_config.get('max_tokens', 50)
            temperature = output_config.get('temperature', 0.7)
            fetch_all = output_config.get('fetch_all', False)

            # Determine whether to fetch based on 'fetch_all' and cell content
            current_value = row_data.get(output_column, None)
            should_fetch = fetch_all or (pd.isna(current_value) or str(current_value).strip() == '')

            if not should_fetch:
                logger.info(f"Skipping '{output_column}' for row {row_number} as it is already populated and 'fetch_all' is False.")
                continue

            # Prepare the prompt by inserting input column values
            prompt = prompt_template
            for input_col in self.input_columns:
                prompt += f"\n{input_col}: {row_data[input_col]}"
            print(prompt)

            # Retry logic
            for attempt in range(1, self.retry_attempts + 1):
                result = self.openai.create_completion(prompt, max_tokens, temperature)
                if result:
                    # Write directly to the cell
                    column_letter = self.get_column_letter(output_column)
                    cell_reference = f"{column_letter}{row_number}"
                    self.sheet[cell_reference].value = result
                    logger.info(f" - {output_column} updated in cell {cell_reference}: {result}")
                    break
                else:
                    logger.warning(f"Attempt {attempt} failed for '{output_column}' in row {row_number}. Retrying in {self.retry_delay} seconds...")
                    time.sleep(self.retry_delay)
            else:
                logger.error(f"All retry attempts failed for '{output_column}' in row {row_number}.")

            # Sleep between API requests to respect rate limits
            time.sleep(self.sleep_time)

    def get_column_letter(self, column_name: str) -> str:
        """
        Retrieves the Excel column letter for a given column name.

        Args:
            column_name (str): The name of the column.

        Returns:
            str: The Excel column letter.

        Raises:
            ValueError: If the column name is not found.
        """
        for idx, cell in enumerate(self.sheet[1], start=1):
            if cell.value == column_name:
                return cell.column_letter
        raise ValueError(f"Column '{column_name}' not found in the Excel sheet.")

    def process_excel(self):
        """
        Processes the entire Excel sheet by iterating through each row and updating cells as needed.
        """
        try:
            for idx, row in enumerate(self.sheet.iter_rows(min_row=2, values_only=True), start=2):
                row_data = pd.Series(row, index=[cell.value for cell in self.sheet[1]])
                self.process_row(idx, row_data)
        except Exception as e:
            logger.error(f"Error processing Excel file: {e}")
            raise
        finally:
            # Save the workbook after processing all rows
            self.save_workbook()

# ---------------------------- Main Execution ---------------------------- #

def main():
    # Define paths
    config_path = Path(__file__).parent.parent / 'config' / 'config.yaml'

    # Load configuration
    config = load_config(str(config_path))

    # Retrieve OpenAI API key environment variable name from configuration
    api_key_env_var = config['openai']['api_key_env_var']
    api_key = os.getenv(api_key_env_var)
    if not api_key:
        logger.error(f"OpenAI API key not found. Please set the '{api_key_env_var}' environment variable.")
        return

    # Retrieve OpenAI model and system message from configuration
    model = config['openai'].get('model', 'gpt-4o-mini')
    system_message = config['openai'].get('system_message', "You are a helpful assistant.")

    # Initialize OpenAI client
    openai_client = OpenAIClient(api_key=api_key, model=model, system_message=system_message)

    # Initialize Excel processor
    processor = ExcelProcessor(config=config, openai_client=openai_client)

    # Process Excel file
    processor.process_excel()

    logger.info("Excel processing completed successfully.")

if __name__ == "__main__":
    main()
