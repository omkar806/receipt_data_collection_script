import requests
import base64
import PyPDF2
from docx import Document
import io
import os


def extract_text_from_pdf(pdf_data: bytes) -> str:
    """Extract text from PDF bytes."""
    text = ""
    with io.BytesIO(pdf_data) as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page_num in range(len(pdf_reader.pages)):
            page = pdf_reader.pages[page_num]
            text += page.extract_text() or ""
    return text


def extract_text_from_docx(docx_data: bytes) -> str:
    """Extract text from DOCX bytes."""
    text = ""
    doc = Document(io.BytesIO(docx_data))
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text


def extract_text_from_attachment(filename: str, data: bytes) -> str:
    """Determine how to extract text based on file extension."""
    if filename.lower().endswith('.pdf'):
        return extract_text_from_pdf(data)
    elif filename.lower().endswith('.docx'):
        return extract_text_from_docx(data)
    else:
        return "Unsupported or no text extraction for this file type."


def save_attachment_locally(filename: str, data: bytes, output_dir: str = "attachments"):
    """
    Save the raw attachment data (bytes) to a local file.
    Creates 'attachments' folder by default if not exists.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    cleaned_filename = filename.replace("/", "")
    filepath = os.path.join(output_dir, cleaned_filename)
    with open(filepath, 'wb') as f:
        f.write(data)
    print(f"Saved attachment: {filepath}")


def fetch_and_save_attachments(access_token: str):
    """
    Fetch messages from Gmail using the provided `access_token`,
    filter by `brand_name`, and save attachments locally.
    """

    # Optionally, verify user info if needed:
    user_info = requests.get(
        "https://www.googleapis.com/oauth2/v1/userinfo",
        headers={"Authorization": f"Bearer {access_token}"}
    )
    if user_info.status_code != 200:
        raise Exception(f"Failed to fetch user info: {user_info.text}")

    # Build Gmail query
    user_query = (
        '(subject:"your order" OR subject:receipts OR subject:receipt OR subject:invoice '
        'OR subject:invoices OR subject:"insurance" OR subject:"health report" '
        'OR category:purchases OR label:receipts OR label:invoices '
        'OR label:insurance OR label:health) has:attachment'
    )

    # Initialize pagination
    page_token = None
    messages = []

    # Paginate through all search results
    while True:
        gmail_url = f"https://www.googleapis.com/gmail/v1/users/me/messages?q={user_query}"
        if page_token:
            gmail_url += f"&pageToken={page_token}"

        response = requests.get(gmail_url, headers={"Authorization": f"Bearer {access_token}"})
        gmail_data = response.json()

        if "messages" in gmail_data:
            messages.extend(gmail_data["messages"])

        if "nextPageToken" in gmail_data:
            page_token = gmail_data["nextPageToken"]
        else:
            break

    print(f"Total messages found: {len(messages)}")

    attachment_count = 0
    # Loop through all messages and fetch attachments
    for i, message in enumerate(messages):
        if not message:
            continue

        message_id = message.get("id")
        if not message_id:
            continue

        # Get full message details
        message_url = f"https://www.googleapis.com/gmail/v1/users/me/messages/{message_id}"
        message_response = requests.get(message_url, headers={"Authorization": f"Bearer {access_token}"})
        message_data = message_response.json()

        # If you need the body text for debugging or inspection:
        if "payload" in message_data:
            payload = message_data["payload"]
            if "body" in payload and "data" in payload["body"]:
                body_data = payload["body"]["data"]
                body_content = base64.urlsafe_b64decode(body_data.encode("UTF-8")).decode("UTF-8")
        # Look for parts that contain attachments
        if "payload" in message_data and "parts" in message_data["payload"]:
            for part in message_data["payload"]["parts"]:
                if "body" in part and "attachmentId" in part["body"]:
                    attachment_id = part["body"]["attachmentId"]
                    filename = part.get("filename", "untitled.txt")

                    # Fetch the attachment data
                    attachment_url = (
                        f"https://www.googleapis.com/gmail/v1/users/me/messages/{message_id}"
                        f"/attachments/{attachment_id}"
                    )
                    attachment_resp = requests.get(
                        attachment_url, headers={"Authorization": f"Bearer {access_token}"}
                    )
                    attachment_data = attachment_resp.json()
                    data_base64 = attachment_data.get("data")
                    if not data_base64:
                        continue

                    # Decode Base64 data
                    attachment_content = base64.urlsafe_b64decode(data_base64)

                    # Save locally
                    save_attachment_locally(filename, attachment_content)

                    # Optionally extract text if PDF or DOCX
                    text = extract_text_from_attachment(filename, attachment_content)
                    print(f"Extracted text from {filename}:")
                    print(text)

                    attachment_count += 1

    print(f"Total attachments processed: {attachment_count}")


def main():
    """
    Entry point for the script.
    You can adapt how you retrieve the access token and brand name (e.g., from command line args).
    """
    # For demonstration, we simply use input() to get the access token and brand name.
    access_token_new = input("Enter your Google OAuth2 Access Token: ").strip()
    brand_name = input("Enter the brand name: ").strip()

    if not access_token_new:
        print("No access token provided. Exiting.")
        return

    try:
        fetch_and_save_attachments(access_token_new)
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
