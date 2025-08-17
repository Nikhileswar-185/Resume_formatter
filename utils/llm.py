"""
LLM (Large Language Model) utility module for generating structured JSON content from a system prompt and extracted resume content.
It uses the Google Gemini API to parse the input and return structured data.
"""


from google import genai
from dotenv import load_dotenv
import os
import json

# Load variables from .env into environment
load_dotenv()

api_key = os.getenv("GEMINI_API_KEY ")

client = genai.Client(api_key=api_key)


def parse_json(system_prompt ,extracted_resume_content):
    
    resp = client.models.generate_content(
        model="gemini-1.5-flash",
        contents=[
        {
            "role": "user",
            "parts": [
                {
                    "text": f"{system_prompt}\n\n{extracted_resume_content}"
                }
            ]
        }
    ],
        config={
            "temperature": 0.7,  # same as OpenAI's temperature
            "response_mime_type": "application/json"  # ensures valid JSON
        }
    )

    # Convert to Python dict
    data = json.loads(resp.text)
    # print(json.dumps(data, indent=2))
    return data