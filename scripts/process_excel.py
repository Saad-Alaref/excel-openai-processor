# scripts/process_excel.py

import os
import time
import logging
import yaml
import pandas as pd
import openai
from dotenv import load_dotenv
from pathlib import Path
from typing import Any, Dict, List, Optional

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
    def __init__(self, api_key: str):
        """
        Initializes the OpenAI client with the provided API key.

        Args:
            api_key (str): OpenAI API key.
        """
        openai.api_key = api_key
        logger.info("OpenAI client initialized.")

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
            response = openai.ChatCompletion.create(
                model="gpt-4",
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=max_tokens,
                temperature=temperature
            )
            content = response.choices[0].message['content'].strip()
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
        self.output_path = config['excel']['output_path']
        self.sheet_name = config['excel'].get('sheet_name', 'Sheet1')
        self.input_columns = config['columns']['input']
        self.output_columns = config['columns']['output']
        self.sleep_time = config['processing'].get('sleep_time', 1)
        self.retry_attempts = config['processing'].get('retry_attempts', 3)
        self.retry_delay = config['processing'].get('retry_delay', 5)

        logger.info(f"ExcelProcessor initialized with input file '{self.input_path}' and output file '{self.output_path}'.")

    def load_excel(self) -> pd.DataFrame:
        """
        Loads the Excel file into a pandas DataFrame.

        Returns:
            pd.DataFrame: The loaded DataFrame.
        """
        try:
            df = pd.read_excel(self.input_path, sheet_name=self.sheet_name, engine='openpyxl')
            logger.info(f"Excel file '{self.input_path}' loaded successfully.")
            return df
        except Exception as e:
            logger.error(f"Failed to load Excel file '{self.input_path}': {e}")
            raise

    def save_excel(self, df: pd.DataFrame):
        """
        Saves the DataFrame to an Excel file.

        Args:
            df (pd.DataFrame): The DataFrame to save.
        """
        try:
            df.to_excel(self.output_path, index=False, engine='openpyxl')
            logger.info(f"Updated Excel file saved to '{self.output_path}'.")
        except Exception as e:
            logger.error(f"Failed to save Excel file '{self.output_path}': {e}")
            raise

    def process_row(self, row: pd.Series) -> Dict[str, Any]:
        """
        Processes a single row by generating outputs using OpenAI.

        Args:
            row (pd.Series): The DataFrame row.

        Returns:
            Dict[str, Any]: Generated outputs for the row.
        """
        outputs = {}
        for output_column, output_config in self.output_columns.items():
            prompt_template = output_config['prompt']
            max_tokens = output_config.get('max_tokens', 50)
            temperature = output_config.get('temperature', 0.7)

            # Prepare the prompt by inserting input column values
            prompt = prompt_template
            for input_col in self.input_columns:
                prompt += f"\n{input_col}: {row[input_col]}"

            # Retry logic
            for attempt in range(1, self.retry_attempts + 1):
                result = self.openai.create_completion(prompt, max_tokens, temperature)
                if result:
                    outputs[output_column] = result
                    logger.debug(f"Generated '{output_column}': {result}")
                    break
                else:
                    logger.warning(f"Attempt {attempt} failed for '{output_column}'. Retrying in {self.retry_delay} seconds...")
                    time.sleep(self.retry_delay)
            else:
                outputs[output_column] = None
                logger.error(f"All retry attempts failed for '{output_column}'.")

            # Sleep between API requests to respect rate limits
            time.sleep(self.sleep_time)

        return outputs

    def process_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Processes the entire DataFrame by applying OpenAI completions.

        Args:
            df (pd.DataFrame): The input DataFrame.

        Returns:
            pd.DataFrame: The updated DataFrame with generated outputs.
        """
        for index, row in df.iterrows():
            logger.info(f"Processing row {index + 1}/{len(df)}: {row[self.input_columns].to_dict()}")
            generated_outputs = self.process_row(row)
            for column, value in generated_outputs.items():
                df.at[index, column] = value
                logger.info(f" - {column}: {value if value else 'Failed to generate'}")
        return df

# ---------------------------- Main Execution ---------------------------- #

def main():
    # Define paths
    config_path = Path(__file__).parent.parent / 'config' / 'config.yaml'

    # Load configuration
    config = load_config(str(config_path))

    # Retrieve OpenAI API key from environment variables
    api_key = os.getenv('OPENAI_API_KEY')
    if not api_key:
        logger.error("OpenAI API key not found. Please set the 'OPENAI_API_KEY' environment variable.")
        return

    # Initialize OpenAI client
    openai_client = OpenAIClient(api_key=api_key)

    # Initialize Excel processor
    processor = ExcelProcessor(config=config, openai_client=openai_client)

    # Load Excel data
    df = processor.load_excel()

    # Process DataFrame
    updated_df = processor.process_dataframe(df)

    # Save updated Excel
    processor.save_excel(updated_df)

    logger.info("Excel processing completed successfully.")

if __name__ == "__main__":
    main()
