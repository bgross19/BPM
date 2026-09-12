# Code Review: BPM Maintenance Portal

This document outlines a code review for the BPM Maintenance Portal (Google Apps Script project). It serves as a foundation for a larger project, so establishing good practices early is crucial.

## 1. Architectural & Structural Improvements

*   **Separation of Concerns (Frontend):** Currently, `Index.html` contains the HTML structure, CSS styles (via Tailwind and a style block), and all JavaScript logic in one large file.
    *   **Recommendation:** Follow standard Google Apps Script practices by separating these into `Index.html`, `Stylesheet.html`, and `JavaScript.html`. Create an `include(filename)` function in `Code.gs` to inject them into `Index.html`.
*   **Missing Assets:** The FontAwesome link in `Index.html` is marked as `[suspicious link removed]`.
    *   **Recommendation:** Restore a standard FontAwesome CDN link so the UI icons render correctly.
*   **Database Schema:** The schema defined in `setupDatabase` is a good start. However, using string names for properties and tasks directly in the logs might cause issues if a property name changes.
    *   **Recommendation (Long-term):** Consider using unique UUIDs for properties and tasks in the reference tables (`Properties`, `Tasks`) and storing those IDs in `fact_Work_Logs` instead of the raw names.

## 2. Backend Performance & Reliability (Apps Script Quotas)

*   **Inefficient Google Sheets Writes (The `setValue` Anti-Pattern):** Functions like `updateWorkLog`, `updateApprovalStatus`, `markPayrollPaid`, and `markPropertyBilled` iterate through rows and call `sheet.getRange(...).setValue(...)` for individual cells. This is an anti-pattern in Apps Script. Each call to the Spreadsheet service is slow. If the sheet gets large, this will cause the script to hit the 6-minute execution timeout limit.
    *   **Recommendation:** Read the entire data range into memory (`getValues()`), modify the 2D array in JavaScript, and write the entire array back in one single operation using `sheet.getRange(...).setValues(modifiedArray)`.
*   **Temporary File Leaks:** The `exportInvoicePDF` function creates a Google Doc to build the invoice, converts it to a PDF, and saves the PDF. However, it never deletes the temporary Google Doc. Over time, the Google Drive will be cluttered with hundreds of temporary invoice docs.
    *   **Recommendation:** Add `DriveApp.getFileById(doc.getId()).setTrashed(true);` after generating the PDF to clean up the temporary document.
*   **Robust Folder Lookup:** `getOrCreateReceiptsFolder` looks up the folder by name. If a user manually renames the folder in Google Drive, the script will create a new, duplicate folder.
    *   **Recommendation:** Once created or found, save the Folder ID in `PropertiesService.getScriptProperties()`. On subsequent runs, look up the folder by ID first. Fall back to name search or creation if the ID is missing or invalid.
*   **Inefficient Data Fetching:** `getAdminPayrollData` fetches the entire `dim_Employees`, `fact_Work_Logs`, and `fact_Material_Expenses` sheets into memory.
    *   **Recommendation (Long-term):** As the dataset grows, this will become slow. Consider implementing pagination on the frontend or using the advanced Google Sheets API or BigQuery if the data scales beyond thousands of rows. For now, it's acceptable but should be monitored.

## 3. Security & Validation

*   **Client-Side Trust:** The frontend sends a payload to `submitWorkLog`. There is minimal validation on the server side to ensure the employee isn't submitting logs for other employees or forging data.
    *   **Recommendation:** Ensure critical fields (like `Employee_ID` and `Timestamp`) are always determined securely on the server-side, overriding anything sent from the client. (The current implementation of `submitWorkLog` does this correctly for `EmployeeId`, which is good).
*   **Error Handling:** Error handling is present (`try...catch`), but the errors passed back to the frontend might expose internal script details.
    *   **Recommendation:** Sanitize error messages sent to the client to avoid leaking implementation details.
*   **File Upload Security:** The script accepts base64 file uploads.
    *   **Recommendation:** Implement server-side validation to ensure the uploaded files are actually images or PDFs and not malicious scripts, and enforce reasonable file size limits before processing the base64 string to avoid memory exhaustion.

## 4. UI / UX

*   **Mock Environment:** The frontend relies heavily on a `mockGoogleScriptRun` object for local testing.
    *   **Recommendation:** Keep this mock environment, but ensure there's a robust build/deployment process so that the mock data doesn't accidentally ship to production if the code is compiled differently in the future.
*   **Feedback:** The toast notification system is good.
    *   **Recommendation:** Ensure loading states are clearly communicated to the user, especially during file uploads which can take several seconds. The current implementation disables the submit button and shows a spinner, which is the correct approach.

## Immediate Action Items for Refactoring:

1.  Split `Index.html` into `Index.html`, `JavaScript.html`, and `Stylesheet.html`.
2.  Restore the FontAwesome CDN link.
3.  Refactor Sheet updates to use bulk `setValues()`.
4.  Fix the Google Doc leak in `exportInvoicePDF()`.
5.  Implement property caching for the receipts folder.