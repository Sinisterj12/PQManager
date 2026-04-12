# PQManager (Printer Queue Manager) 🖨️

<div align="center">
  <img src="https://img.shields.io/github/license/Sinisterj12/PQManager?style=for-the-badge&color=2563eb" alt="License" />
  <img src="https://img.shields.io/github/last-commit/Sinisterj12/PQManager?style=for-the-badge&color=16a34a" alt="Last Commit" />
  <img src="https://img.shields.io/badge/Platform-Windows?style=for-the-badge&color=0f766e" alt="Platform" />
  <img src="https://img.shields.io/badge/Status-Stable%20Utility?style=for-the-badge&color=f59e0b" alt="Status" />
</div>

<p align="center">
  A tray-based Windows utility built to keep broken print queues from disrupting real store operations.
</p>

---

## 🚀 Overview

**PQManager**, short for **Printer Queue Manager**, was built to solve a real customer issue involving a legacy NCR virtual terminal setup where print jobs would continuously stack, stall printing, and create repeated support issues throughout the day and even overnight.

When vendor support did not provide a working solution, this tool was built as a practical fix. It runs quietly in the system tray, monitors the selected printer queue, and clears stuck jobs on a recurring cycle so printing remains stable without constant manual intervention.

> [!IMPORTANT]
> This tool is designed for Windows environments and is intended to run as a tray-based executable for continuous background operation.

---

## ✨ What it is built to do

* Prevent printer queues from filling up and blocking printing
* Reduce repeated support calls caused by stuck print jobs
* Run silently in the system tray after launch
* Automatically monitor and clear the target printer queue
* Provide logs for troubleshooting when needed

---

## 🛠️ Key Features

* 🖨️ **Automatic Queue Monitoring**
* 🔄 **Recurring Cleanup Cycle**
* 🎯 **Targeted Printer Selection**
* 🧩 **System Tray Operation**
* 💾 **Settings Persistence**
* 📊 **Logging for Troubleshooting**

---

## ⚡ Why it matters

This is not a generic print utility.

It was built because a real production environment was failing, the issue kept coming back, and support needed a reliable workaround that could run quietly and keep operations moving without constant intervention.

It has continued doing that job for years without recurring issues.

---

## 📥 Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/Sinisterj12/PQManager.git
   cd PQManager
   ```

2. Create and activate a virtual environment:

   ```bash
   python -m venv venv
   .\venv\Scripts\activate
   ```

3. Install dependencies:

   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ Usage

Launch the compiled executable or run the Python entrypoint directly.

Once started, PQManager:

* minimizes to the system tray
* monitors the selected printer queue
* clears stuck print jobs automatically
* logs activity in the background

---

## ⚙️ Configuration

Settings are stored in:

```text
%APPDATA%\PQManager\settings.json
```

---

## 📜 Logging

Logs are stored in:

```text
logs/printer_queue.log
```

---

## 📄 License

[MIT License](LICENSE)

---

<div align="center">
  <sub>Built by <a href="https://github.com/Sinisterj12">Sinisterj12</a></sub>
</div>
