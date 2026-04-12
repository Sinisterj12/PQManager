# PQManager (Printer Queue Manager) 🖨️

<div align="center">
  <img src="https://img.shields.io/badge/License-MIT-2563eb?style=for-the-badge" alt="License" />
  <img src="https://img.shields.io/github/last-commit/Sinisterj12/PQManager?style=for-the-badge&color=16a34a" alt="Last Commit" />
  <img src="https://img.shields.io/badge/Platform-Windows-0f766e?style=for-the-badge" alt="Platform" />
  <img src="https://img.shields.io/badge/Status-Stable%20Utility-f59e0b?style=for-the-badge" alt="Status" />
</div>

<p align="center">
  A tray-based Windows utility built to keep broken print queues from disrupting real store operations.
</p>

---

### 🚀 Overview
**PQManager**, short for **Printer Queue Manager**, was built to solve a real customer issue involving a legacy NCR virtual terminal setup where print jobs would continuously stack, stall printing, and create repeated support issues throughout the day and overnight.

When vendor support did not provide a working solution, this tool was built as a practical workaround. It runs quietly in the system tray, monitors the selected printer queue, and clears stuck jobs on a recurring cycle so printing stays stable without constant manual intervention.

> [!IMPORTANT]
> This tool is designed for **Windows environments** and is intended to run as a tray-based executable for continuous background operation.

---

### ✨ Key Features
- 🖨️ **Automatic Queue Monitoring:** Watches the selected printer queue and reacts before jobs pile up long enough to stall operations.
- 🔄 **Recurring Cleanup Cycle:** Repeatedly clears stuck jobs so staff do not have to babysit printing.
- 🎯 **Targeted Printer Selection:** Focuses on the problem printer instead of acting like a broad system-wide cleanup tool.
- 🧩 **System Tray Operation:** Runs quietly in the background after launch.
- 💾 **Settings Persistence:** Remembers the chosen printer and saved preferences.
- 📊 **Logging for Troubleshooting:** Keeps useful records when support needs to review behavior.

---

### ⚡ Why It Matters
This is not a generic print utility.

It was built because a real production environment was failing, the issue kept returning, and support needed a reliable fix that could run quietly and keep operations moving without constant intervention.

PQManager has been doing exactly that for years without recurring issues.

---

### 🛠️ Quick Start

#### 📥 Installation
```bash
git clone https://github.com/Sinisterj12/PQManager.git
cd PQManager
python -m venv venv
.\venv\Scripts\activate
pip install -r requirements.txt
```

#### ▶️ Usage
Launch the compiled executable or run the Python entrypoint directly.

| Behavior | What it does |
| :--- | :--- |
| Tray Startup | Minimizes to the system tray after launch |
| Queue Monitoring | Watches the selected printer queue |
| Automatic Cleanup | Clears stuck print jobs automatically |
| Background Logging | Records activity for troubleshooting |

---

### ⚙️ Configuration
Settings are stored in `%APPDATA%\PQManager\settings.json`.

```json
{
  "selected_printer": "Target Printer",
  "monitoring_enabled": true
}
```

---

### 📜 Logging
Logs are stored in `logs/printer_queue.log`.

---

### 📄 License
[MIT License](LICENSE)

---

<div align="center">
  <sub>Built by <a href="https://github.com/Sinisterj12">Sinisterj12</a></sub>
</div>
