# Parametrizer 🤖

> **Internal Tool developed for Betta Group** to streamline the configuration of Call Center Skills and Service Types.

## 📋 Project Overview
**Parametrizer** is a Robotic Process Automation (RPA) solution designed to interact with legacy web platforms (C2X). It automates the complex workflow of navigating through tree-view structures, extracting dynamic IDs, and performing bulk data entry.

The solution was built to replace manual data entry, eliminating human error and reducing configuration time by approximately 90%.
...

## 🚀 Key Features
* **Data-Driven Execution:** Reads configuration rules directly from Excel (`.xlsx`) and JSON files.
* **Smart Navigation:** Capable of traversing dynamic tree structures to locate specific Profiles and Skills.
* **Resilient Interaction:** Implements "blinded" logic using JavaScript injection and Explicit Waits to handle legacy system latency and overlays.
* **Data Extraction & Mapping:** Automatically crawls the system to map internal IDs (Segment, Profile, Environment) to human-readable names.
* **Error Handling:** Includes retry logic and detailed logging to ensure process continuity.

## 🛠️ Tech Stack
* **Language:** Python 3.14
* **Core Library:** Selenium WebDriver (Chrome)
* **Data Processing:** Pandas, OpenPyXL, JSON
* **Environment Management:** Dotenv (`.env` for secure credential management)

## ⚙️ Architecture
The project is divided into two main modules to ensure safety:
1.  **Extraction Bot (Read-Only):** Navigates the system to map existing IDs and validate names against the input file. Generates a "Safe Map" for execution.
2.  **Execution Bot (Write):** Reads the "Safe Map" and performs the data entry/configuration via direct form injection.

## 🔒 Security
Credentials and sensitive environment variables are managed via `.env` files and are excluded from version control.

---
*Developed by Victor Silva.*