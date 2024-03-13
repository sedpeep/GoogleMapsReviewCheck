# Google Maps Review Checker

This script searches for a given business name on Google Maps and checks if a particular user has reviewed it. It utilizes Selenium to automate the process. The script has an expiration date to ensure its relevance.

## Prerequisites

Ensure you have the following installed:
- Python
- Selenium
- Microsoft Edge WebDriver

## Installation

1. Clone this repository to your local machine.
2. Install the required dependencies using pip:
    ```
    pip install -r requirements.txt
    ```

## Usage

1. Ensure that the `list.xlsx` file contains the necessary information:
    - Business name and corresponding user to check for reviews.
    - The `Status` column should contain the link number (prefixed with `#`) and the username.
2. Run the script:
    ```
    python google_maps_review_checker.py
    ```
3. The script will search for the given business name on Google Maps and check if the specified user has reviewed it.
4. The status will be updated in the `list.xlsx` file with the result of the review check.

## Note

- This script utilizes Selenium, which requires the Microsoft Edge WebDriver. Make sure it's installed and configured properly.
- The script has an expiration date to ensure its validity. If expired, it will terminate with a corresponding message.
- Ensure that the `list.xlsx` file is correctly formatted with the necessary information.

## Disclaimer

This script is provided as-is, without any guarantees or warranties. Use it at your own risk.
