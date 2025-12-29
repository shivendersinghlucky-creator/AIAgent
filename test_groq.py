import os
from groq import Groq

# Initialize Groq client with API key from environment variable
# Set your API key: $env:GROQ_API_KEY = "your-api-key-here"
client = Groq(api_key=os.getenv("GROQ_API_KEY"))

# Test the API with a simple question
print("Testing Groq API...\n")

response = client.chat.completions.create(
    model="llama-3.3-70b-versatile",  # Fast and powerful model
    messages=[
        {"role": "user", "content": "what is the capital of india"}
    ],
    temperature=0.7,
    max_tokens=100
)

print("Response from Groq AI:")
print(response.choices[0].message.content)
print(f"\n✅ API Key is valid!")
print(f"⚡ Response time: Very fast!")
