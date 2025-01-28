# Outlook VBA Auto-Notification for Conflicting Meetings

This VBA script automatically sends a notification when someone schedules a meeting that conflicts with your existing calendar events in Outlook. Instead of outright declining the meeting, it emails the organizer advising them to propose a different time. It also skips sending a notification for meeting cancellations (subject lines containing `"Canceled"`).

---

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Setup Instructions](#setup-instructions)
3. [Replace the Email Address](#replace-the-email-address)
4. [How It Works](#how-it-works)
---

## Prerequisites

- **Windows Desktop** version of Outlook (VBA is not supported in Outlook on the web or most current Mac versions).
- Ability to run or enable macros (some corporate environments may block this).
- Basic familiarity with the Visual Basic Editor in Outlook.

---

## Setup Instructions

1. **Enable Macros in Outlook:**
   - In Outlook, go to **File** → **Options** → **Trust Center** → **Trust Center Settings...** → **Macro Settings**.
   - Select either:
     - **Notifications for all macros**; or  
     - **Enable all macros** (less secure; recommended only for testing).
   - Click **OK** to save changes.

2. **Open the VBA Editor:**
   - Close the Outlook **Options** window.
   - Press **Alt + F11** in Outlook to open the Visual Basic for Applications (VBA) editor.  
     - If `Alt + F11` does not work, go to **Developer** tab → **Visual Basic** (enable the Developer tab if needed via **File** → **Options** → **Customize Ribbon**).

3. **Insert or Edit `ThisOutlookSession`:**
   - In the VBA editor’s **Project** pane (usually on the left), locate **`Project1 (VbaProject.OTM)`** → **Microsoft Outlook Objects** → **ThisOutlookSession**.
   - Copy and paste the script from the VBA file (conflictAutoresponder.vba) into the **ThisOutlookSession** code window.

4. **Save and Restart Outlook:**
   - Click the **Save** icon or go to **File** → **Save VBAProject** in the editor.
   - Close the VBA editor.
   - **Close and reopen** Outlook to ensure the script’s startup procedure runs.

---

## Replace the Email Address

Inside the script, you’ll see a placeholder email address (`email@email.com`). Replace it with **your** email (or whichever mailbox the script monitors) in the `SendConflictNotification` subroutine, so the automated messages reference the correct mailbox.

---

## How It Works

The script runs whenever a *new* meeting request arrives in your Inbox:

1. It checks if the meeting subject contains `Canceled` (case-insensitive). If so, it **skips** sending a conflict message (this prevents notifications for canceled meetings).
2. It then retrieves the proposed AppointmentItem to compare its start/end times with your existing calendar items.
3. If there’s a conflict, it sends a condescending, automated message back to the sender, including:
   - The subject of the conflicting meeting.
   - A notice to use Outlook’s Scheduling Assistant to find a better time.
