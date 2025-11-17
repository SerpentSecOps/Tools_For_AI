import os
import argparse
import logging
from openai import OpenAI, APIError, RateLimitError
from dotenv import load_dotenv

def main():
    load_dotenv()
    parser = argparse.ArgumentParser(description="A CLI tool to interact with the DeepSeek API.")
    parser.add_argument("prompt", type=str, help="The prompt to send to the DeepSeek API.")
    args = parser.parse_args()

    logging.basicConfig(filename='deepseek_conversation.log', level=logging.INFO, format='%(asctime)s - %(message)s')

    api_key = os.getenv("DEEPSEEK_API_KEY")
    if not api_key or api_key == "your_api_key_here":
        print("Error: DEEPSEEK_API_KEY not found or not set in .env file.")
        return

    client = OpenAI(
        api_key=api_key,
        base_url="https://api.deepseek.com"
    )

    try:
        logging.info(f"Prompt: {args.prompt}")
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "user", "content": args.prompt}
            ],
            temperature=0.7,
            max_tokens=2000
        )
        response_content = response.choices[0].message.content
        logging.info(f"Response: {response_content}")
        print(response_content)
    except RateLimitError:
        error_message = "Rate limit exceeded. Please slow down requests."
        logging.error(error_message)
        print(error_message)
    except APIError as e:
        error_message = f"API error: {e}"
        logging.error(error_message)
        print(error_message)

if __name__ == "__main__":
    main()
