import argparse
import sys

try:
    from openpyxl import Workbook
except ImportError as exc:
    print("Missing required package: openpyxl")
    print("Install it with: pip install -r requirements.txt")
    raise SystemExit(1) from exc


def save_to_excel(filename: str, data: list[dict], sheet_name: str = "Test Results") -> None:
    """Save a list of dictionaries to an Excel file."""
    if not data:
        raise ValueError("No data to save.")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = sheet_name

    headers = list(data[0].keys())
    worksheet.append(headers)

    for record in data:
        row = [record.get(header, "") for header in headers]
        worksheet.append(row)

    workbook.save(filename)
    print(f"Saved {len(data)} rows to '{filename}'")


def sample_test_data() -> list[dict]:
    return [
        {
            "Test Case ID": "TC-001",
            "Description": "Verify login with valid credentials",
            "Expected Result": "User is logged in successfully",
            "Actual Result": "User is logged in successfully",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "No issues found"
        },
        {
            "Test Case ID": "TC-002",
            "Description": "Verify login with invalid password",
            "Expected Result": "Error shown for invalid credentials",
            "Actual Result": "Error shown for invalid credentials",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Behavior is correct"
        },
        {
            "Test Case ID": "TC-003",
            "Description": "Verify upload feature rejects unsupported format",
            "Expected Result": "Upload fails with proper error",
            "Actual Result": "Upload fails with proper error",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Requires regression test coverage"
        },
        {
            "Test Case ID": "TC-004",
            "Description": "Verify forgot password sends reset email",
            "Expected Result": "Password reset email is sent",
            "Actual Result": "Password reset email is sent",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Email delivery verified"
        },
        {
            "Test Case ID": "TC-005",
            "Description": "Verify profile update saves new display name",
            "Expected Result": "Display name updates successfully",
            "Actual Result": "Display name updates successfully",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Field validation passes"
        },
        {
            "Test Case ID": "TC-006",
            "Description": "Verify profile update rejects invalid phone number",
            "Expected Result": "Error shown for invalid phone number",
            "Actual Result": "Error shown for invalid phone number",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Validation error is clear"
        },
        {
            "Test Case ID": "TC-007",
            "Description": "Verify user can search for existing records",
            "Expected Result": "Relevant records are returned",
            "Actual Result": "Relevant records are returned",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Search works correctly"
        },
        {
            "Test Case ID": "TC-008",
            "Description": "Verify search returns no results for unknown term",
            "Expected Result": "No results message is displayed",
            "Actual Result": "No results message is displayed",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Edge case handled"
        },
        {
            "Test Case ID": "TC-009",
            "Description": "Verify dashboard loads within acceptable time",
            "Expected Result": "Dashboard loads under 3 seconds",
            "Actual Result": "Dashboard loads under 3 seconds",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Performance acceptable"
        },
        {
            "Test Case ID": "TC-010",
            "Description": "Verify logout ends session and redirects to login",
            "Expected Result": "User is logged out and redirected",
            "Actual Result": "User is logged out and redirected",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Session cleared successfully"
        },
        {
            "Test Case ID": "TC-011",
            "Description": "Verify invalid email format is rejected at signup",
            "Expected Result": "Signup shows email format error",
            "Actual Result": "Signup shows email format error",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Input validation enforced"
        },
        {
            "Test Case ID": "TC-012",
            "Description": "Verify password strength requirements are enforced",
            "Expected Result": "Weak passwords are rejected",
            "Actual Result": "Weak passwords are rejected",
            "Status": "Pass",
            "Severity": "High",
            "Notes": "Security requirement validated"
        },
        {
            "Test Case ID": "TC-013",
            "Description": "Verify user role permissions prevent unauthorized access",
            "Expected Result": "Access denied for restricted pages",
            "Actual Result": "Access denied for restricted pages",
            "Status": "Pass",
            "Severity": "High",
            "Notes": "Authorization works"
        },
        {
            "Test Case ID": "TC-014",
            "Description": "Verify file download completes successfully",
            "Expected Result": "File downloads without errors",
            "Actual Result": "File downloads without errors",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Download feature stable"
        },
        {
            "Test Case ID": "TC-015",
            "Description": "Verify user cannot upload file above size limit",
            "Expected Result": "File size error is displayed",
            "Actual Result": "File size error is displayed",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Upload restrictions validated"
        },
        {
            "Test Case ID": "TC-016",
            "Description": "Verify notification settings can be toggled",
            "Expected Result": "Notification preferences save successfully",
            "Actual Result": "Notification preferences save successfully",
            "Status": "Pass",
            "Severity": "Low",
            "Notes": "Settings persistence verified"
        },
        {
            "Test Case ID": "TC-017",
            "Description": "Verify password reset link expires after time limit",
            "Expected Result": "Expired link shows invalid message",
            "Actual Result": "Expired link shows invalid message",
            "Status": "Pass",
            "Severity": "High",
            "Notes": "Security timing validated"
        },
        {
            "Test Case ID": "TC-018",
            "Description": "Verify UI displays error when service is unavailable",
            "Expected Result": "Service unavailable message is shown",
            "Actual Result": "Service unavailable message is shown",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Error handling tested"
        },
        {
            "Test Case ID": "TC-019",
            "Description": "Verify pagination navigates between result pages",
            "Expected Result": "Page content changes on navigation",
            "Actual Result": "Page content changes on navigation",
            "Status": "Pass",
            "Severity": "Medium",
            "Notes": "Pagination is functional"
        },
        {
            "Test Case ID": "TC-020",
            "Description": "Verify session timeout logs user out automatically",
            "Expected Result": "User is logged out after inactivity",
            "Actual Result": "User is logged out after inactivity",
            "Status": "Pass",
            "Severity": "High",
            "Notes": "Session timeout behavior validated"
        },
        {
            "Test Case ID": "TC-021",
            "Description": "Verify two-factor authentication setup process",
            "Expected Result": "2FA setup completes successfully",
            "Actual Result": "Setup fails at QR code validation",
            "Status": "Fail",
            "Severity": "High",
            "Notes": "QR code scanner not working properly"
        },
        {
            "Test Case ID": "TC-022",
            "Description": "Verify API response time under load",
            "Expected Result": "API responds within 500ms",
            "Actual Result": "API responds in 2500ms under load",
            "Status": "Fail",
            "Severity": "High",
            "Notes": "Performance degradation observed"
        },
        {
            "Test Case ID": "TC-023",
            "Description": "Verify email notification on record update",
            "Expected Result": "Email is sent within 1 minute",
            "Actual Result": "Email sent after 15 minutes delay",
            "Status": "Fail",
            "Severity": "Medium",
            "Notes": "Email service experiencing delays"
        },
        {
            "Test Case ID": "TC-024",
            "Description": "Verify export data to CSV format",
            "Expected Result": "CSV file is generated correctly",
            "Actual Result": "CSV generated but encoding issues detected",
            "Status": "Fail",
            "Severity": "Medium",
            "Notes": "UTF-8 encoding problem in CSV"
        },
        {
            "Test Case ID": "TC-025",
            "Description": "Verify mobile app responsiveness on small screens",
            "Expected Result": "UI renders correctly on 320px width",
            "Actual Result": "Buttons overflow on small screens",
            "Status": "Fail",
            "Severity": "High",
            "Notes": "CSS media queries need adjustment"
        }
    ]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Save test data to an Excel file.")
    parser.add_argument(
        "-o",
        "--output",
        default="openfabric_test_results.xlsx",
        help="Output Excel filename"
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    data = sample_test_data()
    save_to_excel(args.output, data)
