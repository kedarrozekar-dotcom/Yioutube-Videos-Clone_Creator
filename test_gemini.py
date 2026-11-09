import os
from google import genai

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

api_key = os.getenv("GEMINI_API_KEY")
if not api_key:
    raise SystemExit("GEMINI_API_KEY is not set. Add it to your .env file (see .env.example).")

client = genai.Client(api_key=api_key)

response = client.models.generate_content(
    model=os.getenv("GEMINI_TEST_MODEL", "gemini-3.8-flash"),
    contents="Say 'Gemini API is ready!'"
)

print(response.text)
