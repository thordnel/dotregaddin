# DOT.REG Class Record Web Add-in

[![Microsoft Store](https://img.shields.io/badge/Available_in-Microsoft_Store-blue.svg)](#) [![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](#)

## Overview

The **DOT.REG Class Record Web Add-in** provides a secure, seamless interface between the host application (Microsoft Office) and the user's dedicated DOT.REG system. It is designed exclusively for consumers equipped with an existing DOT.REG backend configuration.

Can work offline after first sync. Sync only data when required (downloading class updates (new section assigned, udpated trainees, class schedule, offcial dates or submission of grades). User may be required to enter password upon connetcion to the reqistrar for these tasks. Succeeding request wil not require, but only after an hour of last access for your security.

To eliminate sideloading requirements and ensure a streamlined, secure installation process for end-users, this add-in is published directly via the Microsoft Store.

## Authentication and Access

Access to the add-in requires authentication through a password-protected gateway. The system validates credentials against the user's designated Registrar Server before granting access to the main dashboard.

### Demo Environment Configuration

Users can evaluate the add-in using our dedicated demo environment.

> **Note:** The demo server enters a sleep state during periods of inactivity. It may require a brief initialization time (up to a minute) upon the first login attempt.

* **Registrar Server Name:** `demo`
* **Username:** `demo`
* **Password:** `demo`

## Operational Instruction (demo):
1. Select a batch available for the instructor.
2. Inspect or select any of the shown sheets for the selected batch: eg. Attendance, Gradesheet, Midterm, Final, Trainees List
3. 

## Core Architecture & Functions

The architecture utilizes standard Office Web Add-in components to deliver functionality securely and efficiently:

* **Authentication Gateway** (`login.html` / `login.js`): Handles secure user credential verification and establishes the active session.
* **Taskpane Interface** (`taskpane.html` / `taskpane.js`): Operates as the primary side-panel UI within the host application, enabling direct interaction with DOT.REG records.
* **Dashboard** (`dashboard.html` / `dashboard.js`): Provides a centralized view of system status, data summaries, and quick actions post-authentication.
* **Sync Engine** (`syncEngine.js`): Manages the bidirectional data transfer between the host document and the DOT.REG backend, ensuring data integrity.
* **Add-in Commands** (`commands.js`): Executes specific ribbon-based actions without requiring the full taskpane to be open, optimizing the user experience.

* Sync Attendance
* Download Update on Class Records
* Submit Gradesshets using automatic caluction (will use Midterm and Final Term grade entries), using Raw Grades or direct Grade Point entry on the Gradesheet Page.
* Grade Submission will always overwrite the entries. Verify submissions by counter checking the printed official records
*

## Data Privacy and Transparency

To comply with strict data privacy standards and Microsoft Store transparency requirements, the add-in strictly adheres to the following data handling protocols:

* **Data Source:** The add-in interacts *strictly* with the user-defined DOT.REG Registrar Server.
* **Data Usage:** Data retrieved from the server is used solely to populate the taskpane and dashboard. User inputs within the host application synchronize directly back to the specified server.
* **Data Retention:** The add-in operates strictly as a pass-through client. It caches session data locally *only* during active use and comprehensively clears this data upon session termination or logout.
* **Third-Party Sharing:** Zero data is transmitted to external analytics platforms, advertising networks, or unauthorized third-party servers.

## Security


## System Limitations

* **Backend Requirement:** The add-in functions solely as a client interface and holds no utility without an active, properly configured DOT.REG system backend.
* **Demo Latency:** As mentioned, the demo server incurs a startup delay if accessed from a cold/sleep state.



---

*For support or inquiries, please open an issue in this repository.*
