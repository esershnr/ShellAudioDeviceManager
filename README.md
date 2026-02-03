# 🎧 Shell Audio Device Manager

<div align="center">

![PowerShell](https://img.shields.io/badge/PowerShell-%235391FE.svg?style=for-the-badge&logo=powershell&logoColor=white)
![Windows](https://img.shields.io/badge/Windows-0078D6?style=for-the-badge&logo=windows&logoColor=white)
![License](https://img.shields.io/badge/license-MIT-green?style=for-the-badge)

**A powerful, interactive PowerShell utility to manage Windows audio devices directly from your keyboard.**

[Features](#-features) • [Installation](#-installation) • [Interactive Mode](#-interactive-mode) • [Usage](#-usage)

</div>

---

## 🚀 Overview

**Shell Audio Device Manager** solves the frustration of switching audio devices in Windows. Instead of clicking through multiple menus, you can switch your headphones, speakers, or microphones instantly with a single keystroke.

It creates a visual, interactive menu in your terminal that lets you control everything with your keyboard.

## ✨ Features

- **🔎 Interactive Menu:** Visual list of all your input and output devices.
- **⚡ Instant Switching:** Switch active devices instantly using number keys **[1-9]**.
- **🎤 Mic Toggle:** Mute/Unmute your default microphone with **[M]**.
- **🔇 Audio Mute:** Mute/Unmute your audio output with **[K]**.
- **🔊 Audio Feedback:** Hear a confirmation sound when you switch devices or toggle mute.
- **🔌 Plug & Play:** No installation required, runs natively on Windows 10 & 11 via PowerShell.

---

## ⚡ Quick Start (One-Line Run)

You can run this tool entirely from memory without downloading files:

```powershell
irm https://raw.githubusercontent.com/esershnr/ShellAudioDeviceManager/main/ShellAudioDeviceManager.ps1 | iex
```

---

## 📥 Installation

1.  **Clone the repository:**

    ```powershell
    git clone https://github.com/esershnr/ShellAudioDeviceManager.git
    cd ShellAudioDeviceManager
    ```

2.  **Allow script execution (if not already enabled):**
    ```powershell
    Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
    ```

---

## 🎮 Interactive Mode (Recommended)

This is the main way to use the tool. It opens a dashboard in your terminal.

```powershell
.\ShellAudioDeviceManager.ps1
```

### ⌨️ Controls

| Key           | Action                                                               |
| :------------ | :------------------------------------------------------------------- |
| **[1] - [9]** | **Select Device** (Switches immediately)                             |
| **[TAB]**     | **Switch Category** (Toggle between Speakers and Microphones)        |
| **[M]**       | **Mute/Unmute Microphone** (Toggles default mic with audio feedback) |
| **[K]**       | **Mute/Unmute Audio** (Toggles default output device audio)          |
| **[ESC]**     | **Exit**                                                             |

### 🔊 Audio Feedback

- **Switching Device:** Plays "Windows Proximity" sound.
- **Mute Mic:** Plays "Navigation Start" sound.
- **Unmute Mic:** Plays "Unlock" sound.

---

## 🛠️ Compatibility

- **OS:** Windows 10, Windows 11
- **Shell:** PowerShell 5.1 or PowerShell 7+
- **Privileges:** No Admin rights required for standard usage.

## 🤝 Contributing

Created by **[esershnr](https://github.com/esershnr)**.
Contributions, issues, and feature requests are welcome!

---

<div align="center">

Made with ❤️ for the CLI community.

</div>
