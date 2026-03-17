<img width="859" height="606" alt="image" src="https://github.com/user-attachments/assets/020f7db4-92ef-4282-ae03-003f609e1773" />

<img width="823" height="547" alt="image" src="https://github.com/user-attachments/assets/381c4f72-ea55-4f7b-9e56-724f8eae59a3" />

<img width="827" height="659" alt="image" src="https://github.com/user-attachments/assets/2800d15a-ac14-42f9-b7b9-f9799a48268c" />

Here is a professional, de-identified README.md in English, designed to showcase the project's architecture and your technical skills.

eBOM Automation & Synchronization System
This project provides an end-to-end automation solution for Managing electronic Bill of Materials (eBOM). It bridges the gap between cloud-based collaboration tools (Smartsheet/Google Sheets), enterprise Resource Planning (ERP) systems like SAP, and automated reporting services.

System Architecture & Data Flow：
The system operates as a closed-loop data pipeline. It begins by consolidating raw material requirements from cloud sources, executes complex document maintenance in SAP via RPA-style scripts, and concludes by synchronizing execution results back to the source platforms while notifying stakeholders through automated analytics reports.

Cloud Data Orchestration (Google Apps Script)：
The orchestration layer uses Google Apps Script (GAS) to manage data lifecycle and cross-platform communication. It features automated filtering and deduplication logic to ensure only the most recent material records are processed. By leveraging Smartsheet and Google Drive APIs, it creates a seamless bridge between local file storage and high-level project management dashboards.

SAP RPA Maintenance (Python)：
The core execution engine is a Python-based automation suite that interacts with the SAP GUI using the Win32COM interface. It automates the CV01N and CV02N transactions to create document info records and link them to material masters. This module replaces hours of manual data entry with a high-speed, error-free batch processing routine that includes localized transaction logging for audit trails.

Bidirectional Result Synchronization：
A dedicated synchronization module ensures that "offline" processing results are accurately reflected "online." It parses generated CSV result files and uses the Smartsheet API to programmatically toggle status markers (e.g., checkboxes) for specific rows. This guarantees that project managers have a real-time view of maintenance progress without requiring manual status updates.

Automated Stakeholder Reporting：
The system concludes its daily cycle by generating and dispatching an HTML-formatted summary report. This module scans the unified audit log, identifies records processed within the last 24 hours, and formats them into a clean, professional table sent directly to the engineering team. This transparency ensures all stakeholders are aligned on daily throughput and potential exceptions.

Technical Stack：
Languages: Python (Pandas, Win32COM), JavaScript (Google Apps Script).

APIs: Smartsheet API, Google Drive API, Domo Dataset API.

ERP Integration: SAP GUI Scripting (Document Management System).

Environment: Cross-platform integration between local drives and Google Workspace.

<img width="1276" height="440" alt="image" src="https://github.com/user-attachments/assets/71a8ed48-e7ec-4158-b385-5c43b3f4e056" />

