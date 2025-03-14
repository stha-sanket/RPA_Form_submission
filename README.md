# Google Form Filler from Excel

## Description

This Robocorp script automates the process of filling out a Google Form using data from an Excel file. The script reads the data from the Excel file, navigates to the Google Form URL, fills out the form fields based on the Excel data, submits the form, and clicks the "Submit another response" button to prepare for the next entry.

## Prerequisites

*   **Python 3.7+**
*   **Robocorp Framework**: `pip install robocorp-framework`
*   **Required Libraries**:
    *   `rpaframework`
    *   `rpaframework-browser`
    *   `pandas`
    *   `rpaframework-selenium`
*   Install these libraries using `pip install -r requirements.txt`
*   using the Vs-code extension install:
*   ![Requirement Extension](https://github.com/stha-sanket/RPA-Auto-WebsiteScraper/blob/main/requirement-extension.png?raw=true)
  
## Usage

1.  **Run the Robot:**
 After downloading the extension you can see a Run Task above `@task` decorator

## Code Explanation

*   **Robocorp Framework:** Provides the structure for defining and running tasks.
*   **Browser Library (Selenium):** Controls the web browser (Chrome) for navigating the Google Form and interacting with its elements. This implementation leverages the `Selenium` class from `RPA.Browser.Selenium` directly for browser interaction. This is necessary because interacting with some dynamic elements of the form requires more explicit Selenium functionality.
*   **Pandas:** Reads the data from the `data.xlsx` Excel file into a Pandas DataFrame.
*   **Data Iteration:** The script iterates through each row of the DataFrame, extracting the data for each form field.
*   **Form Filling:** The script uses XPath expressions to locate the form fields on the Google Form and fills them with the corresponding data from the Excel file.
*   **Form Submission:** The script clicks the "Submit" button to submit the form.
*   **"Submit Another Response" Handling:** After submitting the form, the script waits for the "Submit another response" link to become visible and clicks it to prepare for the next entry.
*   **Course Selection:** The script uses conditional statements to select a course from a group of radio buttons based on the value in the "Course Interested" column of the Excel file.

## Important Considerations

*   **Google Form's HTML Structure:** The script relies heavily on XPath expressions to locate form elements. Changes to the Google Form's layout or element IDs will likely break the script. You'll need to update the XPath expressions in the script if the Google Form's HTML changes. Use your browser's developer tools to inspect the HTML and find the appropriate XPath expressions.
*   **XPath Fragility:** The script utilizes very specific XPaths to locate elements. It's **strongly recommended** to find more robust and stable locators whenever possible. Consider using element IDs, names, or CSS selectors. XPath locators are often prone to breaking due to minor changes in the Google Forms HTML structure.
*   **Hardcoded URL:** The Google Form URL (`https://forms.gle/qMFvYEouXZFPCtDB7`) is hardcoded in the script. Consider making this configurable, either by using a configuration file or a command-line argument.
*   **Dynamic Form Elements:** Google Forms can dynamically change its HTML structure, making it difficult to scrape consistently. Be prepared to update the XPath locators frequently.
*   **Error Handling:** The script includes minimal error handling. You should add more robust error handling to catch potential exceptions, such as network errors, invalid data in the Excel file, or changes to the Google Form's HTML structure.
*   **Excel File Structure:** The script assumes that the `data.xlsx` file has the correct column headers. Ensure that the column names in the Excel file match the expected names in the script.
*   **`slowmo`**: The line `browser.configure(slowmo=50)` adds a small delay (50 milliseconds) after each browser action. This can make the script easier to debug, but slows down the overall execution. Remove or reduce the `slowmo` value to make the script run faster.
*   **Rate Limiting:** Google Forms may have rate limits. If you are submitting a large number of forms, you may need to add delays between submissions to avoid being blocked.
*   **Selenium Library**:  The script uses the Selenium library via the Selenium class provided by `RPA.Browser.Selenium`. This gives you more fine-grained control over browser interactions, particularly for more complex form elements that might not work seamlessly with Robocorp's default browser library. It's initialized via `browser_lib = Selenium()` and then used in the script. The use of the Selenium Library requires chromedriver.exe to be present in the PATH environment variable.
*   **Data Types:** Explicitly cast datatypes from excel to the appropriate datatype on the Google form
