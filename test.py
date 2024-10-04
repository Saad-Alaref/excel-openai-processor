import json
from openai import OpenAI
import os

# Initialize the OpenAI client
client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))

# Define JSON schema with enum for difficulty
schema = {
    "type": "object",
    "properties": {
        "difficulty": {
            "type": "string",
            "enum": ["easy", "medium", "hard", "insane", "impossible"]
        }
    },
    "required": ["difficulty"]
}

# Send a request with structured output using the new method
response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "You are a task difficulty classifier."},
        {"role": "user", "content": "Classify the difficulty of this task: walking on 1 leg"}
    ],
    functions=[{
        "name": "classify_difficulty",
        "parameters": schema
    }],
    function_call={"name": "classify_difficulty"}
)

# Convert response to a dictionary
response_dict = response.to_dict()

# The arguments field is a JSON string, so we need to parse it
arguments_str = response_dict["choices"][0]["message"]["function_call"]["arguments"]
arguments_dict = json.loads(arguments_str)

# Extract and print the difficulty value
difficulty_value = arguments_dict["difficulty"]
print(difficulty_value)
