
# Smartsheet Update Project

This project automates the management of Smartsheet data, updating entries based on specified conditions and syncing tracking information from UPS and FedEx. The `sql2.py` script executes every 15 minutes using Microsoft Task Scheduler to maintain up-to-date information on tickets and shipments. 

## Project Structure

- `sql2.py`: Main automation script to handle task scheduling, error handling, email notifications, and logging.
- `config2.py`: Contains the core configuration and utility functions for Smartsheet API calls, database interactions, and tracking updates.

## Dependencies

Make sure to install the required dependencies using:

```bash
pip install -r requirements.txt
```

## Environment Variables

To keep sensitive data secure, the following environment variables should be set in a `.env` file in the project root:

```plaintext
GRT_USER=your_granite_user
GRT_PASS=your_granite_password
SMARTSHEET_ACCESS_TOKEN=your_smartsheet_token
UPS_ACCESS_TOKEN=your_ups_token
FEDEX_CLIENT_ID=your_fedex_client_id
FEDEX_CLIENT_SECRET=your_fedex_client_secret
CW_BASE_URL=your_connectwise_url
CW_COMPANY_ID_PROD=your_connectwise_company_id
CW_PUBLIC_KEY=your_connectwise_public_key
CW_PRIVATE_KEY=your_connectwise_private_key
CW_CLIENT_ID=your_connectwise_client_id
```

The environment variables include:

- **GRT_USER**: Granite user for database connection.
- **GRT_PASS**: Password for the Granite database user.
- **SMARTSHEET_ACCESS_TOKEN**: API token for accessing Smartsheet.
- **UPS_ACCESS_TOKEN**: API token for UPS tracking updates.
- **FEDEX_CLIENT_ID**: Client ID for FedEx API authentication.
- **FEDEX_CLIENT_SECRET**: Client Secret for FedEx API authentication.
- **CW_BASE_URL**: Base URL for ConnectWise API.
- **CW_COMPANY_ID_PROD**: ConnectWise Company ID for production environment.
- **CW_PUBLIC_KEY**: Public key for ConnectWise authentication.
- **CW_PRIVATE_KEY**: Private key for ConnectWise authentication.
- **CW_CLIENT_ID**: Client ID for ConnectWise.

## Usage

To run the main script:

```bash
python sql2.py
```

This will execute the main tasks for updating Smartsheet data and tracking information every 15 minutes. Log files will be generated in the `Smartsheet Automation Logfiles` directory, organized by date.

## Scheduled Execution

The `sql2.py` script should be scheduled to run every 15 minutes via Microsoft Task Scheduler to ensure timely updates.

## Logging

Logs are generated in the `Smartsheet Automation Logfiles` directory, with filenames based on the current date. The logs provide details on the automation tasks and error handling.

---

This README provides all necessary information for configuring and running the project.
