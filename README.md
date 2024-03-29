# Email Assistant

## Overview

The Email Assistant is a Python application that integrates with Gmail to read and respond to emails with attachments. It uses OpenAI's GPT-4 language model to understand the context of the email and generate appropriate responses. The application also utilizes the LlamaIndex library to create a vector store index from the attachments, enabling efficient searching and question answering.

## Features

- Read and process unread emails with attachments
- Extract text from various file formats (PDF, Word, Excel)
- Use OCR to extract text from images in PDF files
- Generate intelligent responses based on email content and attachments
- Send responses back to the email sender
- Handle large attachments by chunking and overlapping text
- Logging for monitoring and error handling
- Periodic execution for continuous operation

## Setup

1. Clone the repository to your local machine.
2. Install the required dependencies listed in the `requirements.txt` file.
3. Obtain API keys and credentials for any required services (e.g., Gmail API).
4. Configure the OpenAI API key
5. Set up the environment variables as needed.
6. Run the main script (`main.py`) to start the email processing and response generation.
   
## Usage
1. Run the application using the `periodic_work` function, which will continuously check for new unread emails with attachments and process them accordingly.
2. Customize the prompt messages and persona for response generation in the main script.
3. Run the script and let it handle incoming emails automatically.
4. Monitor the log files (`emailBot.log`) for any errors or issues.

## Real-life Use Cases

1. **Customer Support**: Automate customer support operations by responding to inquiries and processing attachments, such as invoices, receipts, or documentation.
2. **Legal and Compliance**: Streamline document review processes by extracting text from legal documents, contracts, or compliance-related materials attached to emails.
3. **Research and Analysis**: Analyze research papers, reports, or data files sent via email, and provide summaries or insights based on the content.
4. **Human Resources**: Automate the processing of job applications, resumes, and other HR-related documents received via email.
5. **Project Management**: Assist project managers by understanding project-related emails and attachments, providing updates, or answering queries.
6. **Education and Academia**: Facilitate academic assistance by comprehending course materials, assignments, or research papers shared via email.


## Contributing

Contributions to the Email Assistant project are welcome! If you encounter any issues or have suggestions for improvements, please open an issue or submit a pull request on the project's GitHub repository.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgments

- [OpenAI](https://www.openai.com/) for providing the GPT-4 language model
- [LlamaIndex](https://github.com/jezhiggins/llama-index) for the vector store index implementation
- [Google APIs Client Library for Python](https://github.com/googleapis/google-api-python-client) for integrating with the Gmail API
- [PyMuPDF](https://github.com/pymupdf/PyMuPDF) for PDF text extraction
- [python-docx](https://github.com/python-openxml/python-docx) for Word document text extraction
- [pandas](https://github.com/pandas-dev/pandas) for Excel file handling
- [pytesseract](https://github.com/madmaze/pytesseract) for optical character recognition (OCR)

