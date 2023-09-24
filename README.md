# Photo_Integrated_Email_Automation_System

## Overview

The Automated Email Photo Attachment System is a Python-based automation project that enables personalized email communication by attaching preferred photos to individual recipients based on data provided in an Excel dataset. This project efficiently processes recipient information, associates it with specific photo IDs, retrieves corresponding images from a designated local directory, and sends customized emails with the appropriate photo attachments.

## Features

- Automated email generation and delivery with customized photo attachments.
- Integration with Excel datasets for recipient and photo ID information.
- Efficient data manipulation and processing capabilities.
- Robust file handling for seamless photo retrieval and attachment.
- Demonstrates strong skills in automation and data management.

## Technologies Used

- Python
- Pandas (for data manipulation)
- smtplib (for email communication)
- MIME (for email content creation)
- File handling for image retrieval

## Usage

1. Clone the repository to your local machine.

```bash
git clone https://github.com/yourusername/automated-email-photo-system.git

Install the required Python packages.

pip install pandas

Configure your Gmail credentials in the script (sender_email and sender_password variables).

Place your Excel dataset in the project directory or specify the file path in the script (excel_file_path).

Ensure that the photos associated with the photo IDs are located in a local directory specified in the script (photo_file_path).

Run the script.
