# 🔐 SAP Auto Login using VBS Script

This VBS script allows you to automatically log in to SAP without manually entering your credentials every time.

---

## 📌 Prerequisites
- SAP GUI installed and configured
- SAP Logon application open
- Basic access to edit `.vbs` files

---

## 🚀 Setup Instructions

1. **Download the `Main.vbs` file**  
   Rename it to a meaningful name. Recommended format: `<SYSTEM_NAME><CLIENT_ID>.vbs`

2. **Open the file in an editor**  
Use Notepad++ or Notepad.

3. **Update the placeholders in the script**
Replace the following values:

- `SYSTEM_NAME` → Your SAP system name (as shown in SAP Logon)
- `CLIENT_ID` → Your 3-digit client ID
- `USERNAME` → Your SAP username
- `PASSWORD` → Your SAP password

4. **Save the file**

---

## ▶️ How to Run

1. Open **Windows Terminal / Command Prompt**

2. Navigate to the folder containing your `.vbs` file: `cd <path-to-your-vbs-script>`

3. Run the script: `<your-file-name>.vbs`

⚠️ Make sure **SAP Logon is already open** before running the script.

---

## 💡 Pro Tips

- Set your terminal’s default directory to the script folder to avoid typing the path every time.
- Use clear naming for multiple systems (e.g., `DEV100.vbs`, `QAS200.vbs`).

---

## ⚠️ Security Warning

- This script stores your **SAP password in plain text**.  
- Do **NOT** share the file with others.  
- Store it in a secure location and restrict access.

---

## ✅ Summary

This setup helps streamline your SAP login process, saving time and reducing repetitive manual entry.

---
